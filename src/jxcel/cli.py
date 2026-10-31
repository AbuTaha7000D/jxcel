"""Temporary command bootstrap; the jxcel application is not implemented yet."""

import sys


def main() -> int:
    """Report that the application is not implemented and return failure."""
    print(
        "jxcel is not implemented yet. See ROADMAP.md for project status.",
        file=sys.stderr,
    )
    return 2
