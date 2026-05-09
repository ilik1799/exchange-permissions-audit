# Проверка делегированных прав на почтовые ящики Exchange
#
# Алгоритм:
#   1. Подключение к Exchange, загрузка пользователя, групп, ящиков, описаний
#   2. Send-As — через Get-ACL напрямую из AD (без Exchange-сессии, в 4-8x быстрее Get-ADPermission)
#   3. Send on Behalf — через GrantSendOnBehalfTo (уже загружен из Get-Mailbox)
#   4. Full Access — через Start-Job параллельные джобы (каждый создаёт свою PSSession)
#   5. Сбор результатов → JSON в stdout, прогресс в stderr
#
# Вызов из Python: ps_runner.py передаёт параметры через -File
# Вызов вручную: .\check_permissions.ps1 -TargetUser j.doe -Login CORP\j.doe -Password ***

# Подавление предупреждения PSScriptAnalyzer: пароль приходит из Python как строка
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPlainTextForPassword', '')]
param(
    [Parameter(Mandatory = $true)]
    [string]$TargetUser,                # sAMAccountName проверяемого пользователя

    [string]$ExchangeServer = "http://exchange-srv.corp.local/PowerShell",

    [Parameter(Mandatory = $true)]
    [string]$Login,                     # Логин для подключения (DOMAIN\user)

    [Parameter(Mandatory = $true)]
    [string]$Password,                  # Пароль (передаётся из Python через getpass)

    [int]$MaxThreads = 5,               # Параллельных джобов

    [string]$ConfigPath = "",           # Путь к JSON с ExcludedOUs и ExcludedGroups
    [string]$SendAsGuid = "ab721a54-1e2f-11d0-9819-00aa0040529b",  # GUID Send-As
    [switch]$EnableDebug                # Создавать debug-логи джобов
)

# --- Инициализация ---
$ErrorActionPreference = "SilentlyContinue"
$ProgressPreference = "SilentlyContinue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# Загрузка исключений из JSON-файла (передаётся из Python через tempfile)
if ($ConfigPath -and (Test-Path $ConfigPath)) {
    $config = Get-Content $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $ExcludedOUsMail = @($config.ExcludedOUs)
    $ExcludedGroups = @($config.ExcludedGroups)
} else {
    $ExcludedOUsMail = @()
    $ExcludedGroups = @()
}

# Статус пишется в stderr — Python читает его в реальном времени
function Write-Status($msg) {
    [Console]::Error.WriteLine("[PS] $msg")
}

# ═══════════════════════════════════════════════════════════════════════════════
# 1. ПОДКЛЮЧЕНИЕ К EXCHANGE
# ═══════════════════════════════════════════════════════════════════════════════
Write-Status "Подключение к Exchange..."
$secPass = ConvertTo-SecureString $Password -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential($Login, $secPass)
try {
    $mainSession = New-PSSession -ConfigurationName Microsoft.Exchange `
        -ConnectionUri $ExchangeServer -Authentication Kerberos -Credential $Cred
    Import-PSSession $mainSession -AllowClobber -DisableNameChecking | Out-Null
    Write-Status "Сессия создана"
} catch {
    $errMsg = $_.Exception.Message.Replace('"', '\"')
    Write-Output ("{`"error`": `"Exchange connection failed: $errMsg`"}")
    exit 1
}

# ═══════════════════════════════════════════════════════════════════════════════
# 2. ЦЕЛЕВОЙ ПОЛЬЗОВАТЕЛЬ И ЕГО ГРУППЫ
# ═══════════════════════════════════════════════════════════════════════════════
Write-Status "Загрузка пользователя $TargetUser..."
$user = Get-ADUser -Identity $TargetUser -Properties MemberOf, SamAccountName
if (-not $user) {
    Write-Output ("{`"error`": `"User not found: $TargetUser`"}")
    exit 1
}

# HashSet для O(1) поиска учётных записей (вместо массива + -contains = O(n))
$accountNamesSet = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)
[void]$accountNamesSet.Add("$env:USERDOMAIN\$($user.SamAccountName)")

# DN групп для проверки Send on Behalf
$userAndGroupDNs = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)
[void]$userAndGroupDNs.Add($user.DistinguishedName)

# Добавляем группы пользователя (с фильтрацией исключённых)
foreach ($groupDN in $user.MemberOf) {
    try {
        $grp = Get-ADGroup $groupDN -Properties SamAccountName
        $skip = $false
        foreach ($eg in $ExcludedGroups) {
            if ($grp.Name -like "*$eg*") { $skip = $true; break }
        }
        if ($skip) { continue }
        [void]$accountNamesSet.Add("$env:USERDOMAIN\$($grp.SamAccountName)")
        [void]$userAndGroupDNs.Add($groupDN)
    } catch {}
}

$accountNamesArray = @($accountNamesSet)
Write-Status "Учёток: $($accountNamesSet.Count)"

# ═══════════════════════════════════════════════════════════════════════════════
# 3. ЗАГРУЗКА MAILBOX'ОВ (UserMailbox + SharedMailbox + RoomMailbox + Discovery)
# ═══════════════════════════════════════════════════════════════════════════════
Write-Status "Загрузка почтовых ящиков..."
$allMailboxes = Get-Mailbox -ResultSize Unlimited
$mailboxes = @()
foreach ($mbx in $allMailboxes) {
    $dn = $mbx.DistinguishedName
    $excluded = $false
    foreach ($ou in $ExcludedOUsMail) {
        if ($dn -like "*$ou*") { $excluded = $true; break }
    }
    if (-not $excluded) { $mailboxes += $mbx }
}
$total = $mailboxes.Count
Write-Status "Ящиков: $total из $($allMailboxes.Count)"

# Кешируем данные ящиков для быстрого доступа по Identity
$mailboxIdentities = @()
$mailboxInfoMap = @{}
foreach ($mbx in $mailboxes) {
    $id = $mbx.Identity.ToString()
    $mailboxIdentities += $id

    # GrantSendOnBehalfTo сохраняем сразу — потом не нужен Get-Mailbox
    $sob = @()
    if ($mbx.GrantSendOnBehalfTo) {
        $sob = @($mbx.GrantSendOnBehalfTo | ForEach-Object { $_.ToString() })
    }

    $mailboxInfoMap[$id] = @{
        PrimarySmtpAddress  = $mbx.PrimarySmtpAddress.ToString()
        UserPrincipalName   = $mbx.UserPrincipalName
        DistinguishedName   = $mbx.DistinguishedName
        GrantSendOnBehalfTo = $sob
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# 4. ПРЕДЗАГРУЗКА DESCRIPTION (1 LDAP-запрос вместо N в цикле)
# ═══════════════════════════════════════════════════════════════════════════════
Write-Status "Предзагрузка описаний..."
$descriptionMap = @{}
try {
    Get-ADUser -Filter 'msExchMailboxGuid -like "*"' -Properties UserPrincipalName, Description |
        ForEach-Object {
            if ($_.UserPrincipalName) {
                $descriptionMap[$_.UserPrincipalName] = [string]$_.Description
            }
        }
} catch {}
Write-Status "Описаний: $($descriptionMap.Count)"

# Закрываем основную сессию — дальше Exchange не нужен для Send-As
Remove-PSSession $mainSession -ErrorAction SilentlyContinue

# ═══════════════════════════════════════════════════════════════════════════════
# 5. SEND-AS ЧЕРЕЗ GET-ACL (напрямую из AD, без Exchange-сессии)
#    Get-ACL читает ACL объекта в AD → фильтр по GUID Send-As
#    В 4-8x быстрее Get-ADPermission
# ═══════════════════════════════════════════════════════════════════════════════
Write-Status "Проверка Send-As через Get-ACL..."

Import-Module ActiveDirectory -ErrorAction SilentlyContinue
$currentLocation = Get-Location
Set-Location AD:  # Переключаемся в провайдер AD для работы Get-ACL

$allResults = [System.Collections.Generic.List[PSObject]]::new()  # Generic List вместо += (O(1))
$sendAsCount = 0
$saProgress = 0
$saTimer = [System.Diagnostics.Stopwatch]::StartNew()

foreach ($mbxId in $mailboxIdentities) {
    $saProgress++
    # Прогресс каждые 25 ящиков
    if ($saProgress % 25 -eq 0 -or $saProgress -eq $total) {
        $pct = [math]::Round(($saProgress / $total) * 100)
        $elapsed = $saTimer.Elapsed.ToString('mm\:ss')
        if ($saProgress -gt 0) {
            $secPerItem = $saTimer.Elapsed.TotalSeconds / $saProgress
            $remaining = [math]::Round(($total - $saProgress) * $secPerItem)
            $remMin = [math]::Floor($remaining / 60)
            $remSec = $remaining % 60
            Write-Status "[Send-As] $saProgress/$total ($pct%) | $elapsed | ~${remMin}m${remSec}s | found: $sendAsCount"
        }
    }
    $info = $mailboxInfoMap[$mbxId]
    if (-not $info) { continue }

    $dn = $info.DistinguishedName
    $mailAttr = $info.PrimarySmtpAddress
    $descAttr = ""
    if ($info.UserPrincipalName -and $descriptionMap.ContainsKey($info.UserPrincipalName)) {
        $descAttr = $descriptionMap[$info.UserPrincipalName]
    }

    try {
        # Фильтр: ExtendedRight + GUID Send-As + не унаследованное + наш пользователь/группа
        $acl = (Get-ACL $dn).Access | Where-Object {
            $_.ActiveDirectoryRights -eq "ExtendedRight" -and
            $_.ObjectType -eq $SendAsGuid -and
            $_.AccessControlType -eq "Allow" -and
            -not $_.IsInherited -and
            $accountNamesSet.Contains($_.IdentityReference.ToString())
        }

        foreach ($ace in $acl) {
            $allResults.Add([PSCustomObject]@{
                Mailbox     = $mailAttr
                Description = $descAttr
                GrantedTo   = $ace.IdentityReference.ToString()
                Permission  = "Отправить как"
                Detail      = "Разрешение позволяет делегировать отправку электронной почты от имени владельца."
            })
            $sendAsCount++
        }
    } catch {}
}

Set-Location $currentLocation
$saElapsed = $saTimer.Elapsed.ToString('mm\:ss')
Write-Status "[Send-As] Завершено за $saElapsed. Найдено: $sendAsCount"

# ═══════════════════════════════════════════════════════════════════════════════
# 6. SEND ON BEHALF (в основной сессии, данные уже загружены)
#    GrantSendOnBehalfTo сохранён в mailboxInfoMap из Get-Mailbox
#    Только Get-Recipient для резолва DN → sAMAccountName
# ═══════════════════════════════════════════════════════════════════════════════
Write-Status "Переподключение к Exchange для Send on Behalf..."
$mainSession2 = New-PSSession -ConfigurationName Microsoft.Exchange `
    -ConnectionUri $ExchangeServer -Authentication Kerberos -Credential $Cred
Import-PSSession $mainSession2 -AllowClobber -DisableNameChecking | Out-Null

$sobCount = 0
$sobTimer = [System.Diagnostics.Stopwatch]::StartNew()
$sobProgress = 0

foreach ($mbxId in $mailboxIdentities) {
    $sobProgress++
    if ($sobProgress % 100 -eq 0 -or $sobProgress -eq $total) {
        Write-Status "[SendOnBehalf] $sobProgress/$total"
    }

    $info = $mailboxInfoMap[$mbxId]
    if (-not $info) { continue }
    # Пропускаем ящики без делегатов (большинство) — мгновенно
    if (-not $info.GrantSendOnBehalfTo -or $info.GrantSendOnBehalfTo.Count -eq 0) { continue }

    $mailAttr = $info.PrimarySmtpAddress
    $descAttr = ""
    if ($info.UserPrincipalName -and $descriptionMap.ContainsKey($info.UserPrincipalName)) {
        $descAttr = $descriptionMap[$info.UserPrincipalName]
    }

    foreach ($delegate in $info.GrantSendOnBehalfTo) {
        try {
            $delegateUser = Get-Recipient -Identity $delegate -ErrorAction Stop
            $fullName = "$env:USERDOMAIN\$($delegateUser.SamAccountName)"
            if ($accountNamesSet.Contains($fullName)) {
                $allResults.Add([PSCustomObject]@{
                    Mailbox     = $mailAttr
                    Description = $descAttr
                    GrantedTo   = $fullName
                    Permission  = "Отправить от имени"
                    Detail      = "Разрешение позволяет отправлять письма от имени владельца почтового ящика."
                })
                $sobCount++
            }
        } catch {}
    }
}

Remove-PSSession $mainSession2 -ErrorAction SilentlyContinue
$sobElapsed = $sobTimer.Elapsed.ToString('mm\:ss')
Write-Status "[SendOnBehalf] Завершено за $sobElapsed. Найдено: $sobCount"

# ═══════════════════════════════════════════════════════════════════════════════
# 7. FULL ACCESS ЧЕРЕЗ ПАРАЛЛЕЛЬНЫЕ ДЖОБЫ (Start-Job)
#    Каждый джоб создаёт свою PSSession к Exchange (full language mode)
#    Массивы передаются как JSON-строки (PS расплющивает массивы в -ArgumentList)
#    Invoke-Command не подходит — restricted language mode в remote session
# ═══════════════════════════════════════════════════════════════════════════════
Write-Status "Запуск $MaxThreads параллельных джобов для Full Access..."
$faTimer = [System.Diagnostics.Stopwatch]::StartNew()

# Разбиваем ящики на батчи
$batchSize = [math]::Ceiling($total / $MaxThreads)
$batches = @()
for ($i = 0; $i -lt $total; $i += $batchSize) {
    $end = [math]::Min($i + $batchSize - 1, $total - 1)
    $batches += ,@($mailboxIdentities[$i..$end])
}
Write-Status "Батчей: $($batches.Count) по ~$batchSize ящиков"

# Debug-логи создаются только с флагом --debug
$debugDir = ""
if ($EnableDebug) {
    $debugDir = Join-Path $PSScriptRoot "debug_logs"
    if (-not (Test-Path $debugDir)) {
        New-Item -ItemType Directory -Path $debugDir -Force | Out-Null
    }
    Write-Status "Debug-логи включены: $debugDir"
}

# ScriptBlock для каждого джоба — создаёт свою сессию, проверяет только Get-MailboxPermission
$jobScript = {
    param(
        [string]$BatchIdsJson,      # JSON-массив Identity ящиков
        [string]$AccNamesJson,      # JSON-массив учётных записей для фильтрации
        [string]$ExchServer,
        [string]$ExchLogin,
        [string]$ExchPwd,
        [string]$DebugFile          # Путь к лог-файлу джоба
    )

    # Десериализация массивов из JSON (PS расплющивает массивы в -ArgumentList)
    $BatchIds = $BatchIdsJson | ConvertFrom-Json
    $AccNames = $AccNamesJson | ConvertFrom-Json

    $log = @()
    $log += "START: $(Get-Date -Format 'HH:mm:ss') BatchSize=$($BatchIds.Count) AccNames=$($AccNames.Count)"

    # Подключение к Exchange (каждый джоб — отдельная сессия)
    try {
        $secP = ConvertTo-SecureString $ExchPwd -AsPlainText -Force
        $crd = New-Object System.Management.Automation.PSCredential($ExchLogin, $secP)
        $sess = New-PSSession -ConfigurationName Microsoft.Exchange `
            -ConnectionUri $ExchServer -Authentication Kerberos -Credential $crd
        Import-PSSession $sess -AllowClobber -DisableNameChecking | Out-Null
        $log += "SESSION: OK"
    } catch {
        $log += "SESSION FAILED: $_"
        if ($DebugFile) { $log | Out-File $DebugFile -Encoding UTF8 }
        return @()
    }

    # HashSet для O(1) поиска
    $namesSet = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    foreach ($n in $AccNames) { [void]$namesSet.Add($n) }

    $res = [System.Collections.Generic.List[PSObject]]::new()
    $processed = 0
    $faFound = 0
    $errors = 0

    # Только Get-MailboxPermission — никаких лишних вызовов
    foreach ($mbxId in $BatchIds) {
        $processed++
        try {
            $perms = Get-MailboxPermission -Identity $mbxId -ErrorAction Stop
            foreach ($perm in $perms) {
                if ($perm.User -and
                    $namesSet.Contains($perm.User.ToString()) -and
                    -not $perm.IsInherited -and
                    ($perm.AccessRights -contains "FullAccess")) {
                    $res.Add([PSCustomObject]@{
                        Mailbox    = $mbxId
                        GrantedTo  = $perm.User.ToString()
                        Permission = "FullAccess"
                    })
                    $faFound++
                }
            }
        } catch {
            $errors++
            if ($errors -le 3) {
                $log += "ERR ${mbxId}: $_"
            }
        }
    }

    Remove-PSSession $sess -ErrorAction SilentlyContinue
    $log += "DONE: processed=$processed FA=$faFound errors=$errors"
    if ($DebugFile) { $log | Out-File $DebugFile -Encoding UTF8 }

    return @($res)
}

# Запускаем джобы
$jobs = @()
for ($b = 0; $b -lt $batches.Count; $b++) {
    $debugFile = ""
    if ($EnableDebug) {
        $debugFile = Join-Path $debugDir "job_$($b + 1).log"
    }
    $batchJson = $batches[$b] | ConvertTo-Json -Compress
    $accNamesJson = $accountNamesArray | ConvertTo-Json -Compress
    $job = Start-Job -ScriptBlock $jobScript -ArgumentList @(
        $batchJson,
        $accNamesJson,
        $ExchangeServer,
        $Login,
        $Password,
        $debugFile
    )
    $jobs += @{ Job = $job; Batch = $b + 1; DebugFile = $debugFile }
    Write-Status "Джоб $($b + 1)/$($batches.Count) запущен"
}

Write-Status "Ожидание завершения джобов..."

# Сбор результатов — джобы возвращают Identity, мы подставляем PrimarySmtpAddress и Description
$faCount = 0

foreach ($j in $jobs) {
    try {
        $result = Receive-Job -Job $j.Job -Wait -ErrorAction Stop
        Remove-Job -Job $j.Job -Force -ErrorAction SilentlyContinue
        if ($result) {
            foreach ($item in $result) {
                $mbxId = [string]$item.Mailbox

                # Подстановка PrimarySmtpAddress и Description из кеша
                $mailAttr = $mbxId
                $descAttr = ""
                if ($mailboxInfoMap.ContainsKey($mbxId)) {
                    $info = $mailboxInfoMap[$mbxId]
                    $mailAttr = $info.PrimarySmtpAddress
                    if ($info.UserPrincipalName -and $descriptionMap.ContainsKey($info.UserPrincipalName)) {
                        $descAttr = $descriptionMap[$info.UserPrincipalName]
                    }
                }

                $allResults.Add([PSCustomObject]@{
                    Mailbox     = $mailAttr
                    Description = $descAttr
                    GrantedTo   = [string]$item.GrantedTo
                    Permission  = "Полный доступ"
                    Detail      = "Разрешение позволяет открывать почтовый ящик, просматривать и удалять письма."
                })
                $faCount++
            }
        }

        # Debug-лог джоба (только с --debug)
        if ($j.DebugFile -and (Test-Path $j.DebugFile)) {
            $debugContent = Get-Content $j.DebugFile -Raw
            Write-Status "Батч $($j.Batch): $debugContent"
        }
    } catch {
        Write-Status "Ошибка в батче $($j.Batch): $_"
        if ($j.DebugFile -and (Test-Path $j.DebugFile)) {
            Write-Status "Debug: $(Get-Content $j.DebugFile -Raw)"
        }
    }
}

$faElapsed = $faTimer.Elapsed.ToString('mm\:ss')
Write-Status "[FullAccess] Завершено за $faElapsed. Найдено: $faCount"
if ($EnableDebug) {
    Write-Status "Debug-логи: $debugDir"
}

# ═══════════════════════════════════════════════════════════════════════════════
# 8. JSON-ВЫВОД (Python читает stdout, парсит JSON)
# ═══════════════════════════════════════════════════════════════════════════════
$stopwatch.Stop()
$elapsed = $stopwatch.Elapsed.ToString('mm\:ss')
Write-Status "Готово за $elapsed. Найдено прав: $($allResults.Count)"

$output = @{
    user    = $TargetUser
    total   = $total
    found   = $allResults.Count
    threads = $MaxThreads
    elapsed = $elapsed
    report  = @($allResults)
}
Write-Output ($output | ConvertTo-Json -Depth 5 -Compress)
