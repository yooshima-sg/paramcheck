import argparse

from .checks import CHECKS


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="datacheck",
        description="データチェックプログラム集",
    )
    subparsers = parser.add_subparsers(dest="check", metavar="<check>")
    for module in CHECKS:
        sub = subparsers.add_parser(module.NAME, help=module.DESCRIPTION)
        module.add_arguments(sub)
        sub.set_defaults(_run=module.run)
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    if not getattr(args, "check", None):
        parser.print_help()
        raise SystemExit(1)
    raise SystemExit(args._run(args))
