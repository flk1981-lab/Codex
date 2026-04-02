#!/usr/bin/env python3

import argparse
import json
import os
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Save SMTP settings for TB report email automation.")
    parser.add_argument("--sender-email", required=True)
    parser.add_argument("--recipient-email", required=True)
    parser.add_argument("--smtp-password", required=True)
    parser.add_argument("--smtp-host", default="smtp.qq.com")
    parser.add_argument("--port", type=int, default=465)
    parser.add_argument("--smtp-username")
    parser.add_argument("--no-ssl", action="store_true")
    parser.add_argument("--config-path", default=str(Path.home() / ".codex" / "secrets" / "tb-report-mail.json"))
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    config_path = Path(args.config_path).expanduser().resolve()
    config_path.parent.mkdir(parents=True, exist_ok=True)

    payload = {
        "sender_email": args.sender_email,
        "recipient_email": args.recipient_email,
        "smtp_host": args.smtp_host,
        "port": args.port,
        "use_ssl": not args.no_ssl,
        "smtp_username": args.smtp_username or args.sender_email,
        "smtp_password": args.smtp_password,
    }

    with config_path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
        handle.write("\n")

    os.chmod(config_path, 0o600)
    print(f"Saved mail config to {config_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
