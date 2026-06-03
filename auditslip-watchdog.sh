#!/usr/bin/env bash
set -euo pipefail
exec /usr/bin/python3 /root/projects/auditslip/auditslip_watchdog.py "$@"
