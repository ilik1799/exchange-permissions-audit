"""
Проверка делегированных прав Exchange.
CLI, запуск PowerShell, формирование Excel.

Оптимизации:
  - Send-As: Get-ACL вместо Get-ADPermission
  - Full Access: параллельные Start-Job
  - Send on Behalf: GrantSendOnBehalfTo
  - HashSet для O(1) поиска
  - Предзагрузка Description

Использование:
    py main.py -u j.doe -l CORP\\j.doe
    py main.py -u j.doe -l CORP\\j.doe -t 5 -o report.xlsx -d

"""

import argparse
import getpass
import logging
import sys
from datetime import datetime

import excel_report
import ps_runner
from config import (
    DEFAULT_THREADS,
    EXCHANGE_SERVER,
    EXCLUDED_GROUPS,
    EXCLUDED_OUS,
    SEND_AS_GUID,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Проверка прав на ящики Exchange",
    )
    parser.add_argument(
        "-u",
        "--user",
        required=True,
        help="sAMAccountName",
    )
    parser.add_argument(
        "-l",
        "--login",
        required=True,
        help="Логин (DOMAIN\\user)",
    )
    parser.add_argument(
        "-s",
        "--exchange-server",
        default=EXCHANGE_SERVER,
    )
    parser.add_argument(
        "-t",
        "--threads",
        type=int,
        default=DEFAULT_THREADS,
        help="Джобов (1-18)",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="Путь к .xlsx",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
    )
    parser.add_argument(
        "-d",
        "--debug",
        action="store_true",
        help="Debug-логи в scripts/debug_logs/",
    )
    args = parser.parse_args()

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    password = getpass.getpass(
        f"Пароль для {args.login}: ",
    )

    result = ps_runner.run(
        target_user=args.user,
        login=args.login,
        password=password,
        exchange_server=args.exchange_server,
        max_threads=args.threads,
        excluded_ous=EXCLUDED_OUS,
        excluded_groups=EXCLUDED_GROUPS,
        send_as_guid=SEND_AS_GUID,
        debug=args.debug,
    )

    if result.get("error"):
        log.error("Ошибка: %s", result["error"])
        if "raw" in result:
            log.error("Raw: %s", result["raw"])
        sys.exit(1)

    report = result.get("report", [])
    total = result.get("total", 0)
    threads = result.get("threads", args.threads)
    elapsed = result.get("elapsed", "??:??")

    print(f"\n{'=' * 70}")
    header = "  Делегированные права — почтовые ящики Exchange"
    print(header)
    print(f"  Пользователь: {args.user}")
    prt = len(report)
    print(f"  Ящиков: {total} | Прав: {prt} | Потоков: {threads}")
    print(f"  Время: {elapsed}")
    print(f"  {datetime.now():%d.%m.%Y %H:%M}")
    print(f"{'=' * 70}")

    if not report:
        print("  Делегированных прав не найдено.")
        return

    # !! = Full Access, ! = Send-As, ~ = Send on Behalf
    for i, r in enumerate(report, 1):
        perm = r.get("Permission", "?")
        marker = {
            "Полный доступ": "!!",
            "Отправить как": "!",
            "Отправить от имени": "~",
        }.get(perm, " ")
        mbx = r.get("Mailbox", "?")
        gto = r.get("GrantedTo", "?")
        print(f"  {i}. [{marker}] [{perm}] {mbx} -> {gto}")
    print(f"{'=' * 70}\n")

    output_path = (
        args.output
        or f"exchange_perms_{args.user}_{datetime.now():%Y%m%d_%H%M}.xlsx"
    )
    excel_report.export(
        report,
        args.user,
        total,
        threads,
        elapsed,
        output_path,
    )
    log.info(
        "Отчёт: %s (%d записей)",
        output_path,
        prt,
    )


if __name__ == "__main__":
    main()
