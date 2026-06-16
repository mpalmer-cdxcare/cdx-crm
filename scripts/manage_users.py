#!/usr/bin/env python3
import argparse
import getpass
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app.server import (  # noqa: E402
    APP_STATE_DB_PATH,
    ensure_app_state_db,
    list_local_users,
    set_local_user_admin,
    set_local_user_password,
    upsert_local_user,
)


def prompt_password(username):
    while True:
        password = getpass.getpass(f"Password for {username}: ")
        confirmation = getpass.getpass(f"Confirm password for {username}: ")
        if not password:
            print("Password cannot be blank.", file=sys.stderr)
            continue
        if password != confirmation:
            print("Passwords did not match. Try again.", file=sys.stderr)
            continue
        return password


def cmd_seed(_args):
    ensure_app_state_db()
    seeded = [("michael", "Michael", True), ("emily", "Emily", False)]
    for username, display_name, is_admin in seeded:
        password = prompt_password(username)
        upsert_local_user(username, password, display_name=display_name, is_admin=is_admin)
        suffix = " [admin]" if is_admin else ""
        print(f"Seeded user: {display_name} ({username}){suffix}")
    print(f"App state database: {APP_STATE_DB_PATH}")


def cmd_set_password(args):
    ensure_app_state_db()
    password = prompt_password(args.username)
    set_local_user_password(args.username, password)
    print(f"Updated password for {args.username}")


def cmd_create(args):
    ensure_app_state_db()
    password = prompt_password(args.username)
    upsert_local_user(args.username, password, display_name=args.display_name or args.username)
    print(f"Created or updated user: {args.username}")


def cmd_list(_args):
    ensure_app_state_db()
    rows = list_local_users()
    if not rows:
        print("No local users found.")
        return
    for row in rows:
        status = "active" if row["is_active"] else "inactive"
        role = "admin" if row["is_admin"] else "user"
        force = "must-change-password" if row["must_change_password"] else "normal"
        print(f'{row["username"]}\t{row["display_name"]}\t{role}\t{status}\t{force}\tupdated {row["updated_at"]}')


def cmd_grant_admin(args):
    ensure_app_state_db()
    set_local_user_admin(args.username, True)
    print(f"Granted admin access to {args.username}")


def build_parser():
    parser = argparse.ArgumentParser(
        description="Manage local CDX CRM users and passwords."
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    seed = subparsers.add_parser(
        "seed", help="Create or update the initial Michael and Emily users."
    )
    seed.set_defaults(func=cmd_seed)

    create = subparsers.add_parser(
        "create", help="Create a new local user or update an existing one."
    )
    create.add_argument("username")
    create.add_argument("--display-name", default="")
    create.set_defaults(func=cmd_create)

    set_password = subparsers.add_parser(
        "set-password", help="Reset the password for an existing local user."
    )
    set_password.add_argument("username")
    set_password.set_defaults(func=cmd_set_password)

    list_command = subparsers.add_parser("list", help="List configured local users.")
    list_command.set_defaults(func=cmd_list)

    grant_admin = subparsers.add_parser(
        "grant-admin", help="Grant admin access to an existing local user."
    )
    grant_admin.add_argument("username")
    grant_admin.set_defaults(func=cmd_grant_admin)

    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
