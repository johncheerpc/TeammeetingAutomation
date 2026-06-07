"""
================================================================================
 RUNNER (the local "driver")  ->  shared/runner.py
================================================================================

WHY THIS MODULE EXISTS
----------------------
`main.py` should do as little as possible - just start the program. All the
"how do we run locally" logic (reading command-line arguments, setting up
logging, deciding where the message comes from, and then calling the pipeline)
lives here instead. That keeps main.py tiny and keeps each concern in its own
module.

This module does NOT contain the business logic - that is still in
`shared/processor.py`. This is only the glue that feeds input into it.
================================================================================
"""

import argparse
import json
import logging
import sys
from pathlib import Path

from shared.processor import process_message


def configure_logging(verbose: bool) -> None:
    """Set up logging so pipeline messages appear in your terminal.

    Azure captured logs automatically; locally we configure it ourselves,
    otherwise every `logging.info(...)` inside the pipeline would be silent.

    Args:
        verbose: if True, also show DEBUG-level (more detailed) messages.
    """
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s  %(levelname)-7s  %(message)s",
        datefmt="%H:%M:%S",
    )


def build_arg_parser() -> argparse.ArgumentParser:
    """Describe the command-line options (argparse also builds --help)."""
    parser = argparse.ArgumentParser(
        description="Process a Teams transcript queue message locally.",
    )
    parser.add_argument(
        "path", nargs="?", help="Path to a JSON file containing the queue message."
    )
    parser.add_argument("--file", help="Path to the queue message JSON file.")
    parser.add_argument("--message", help="The queue message JSON as an inline string.")
    parser.add_argument(
        "--verbose", "-v", action="store_true", help="Show extra DEBUG logging."
    )
    return parser


def read_message(args: argparse.Namespace) -> str:
    """Decide WHERE the queue message comes from and return its text.

    Priority order:
        1. --message "<json>"   (inline string)
        2. file path            (--file or the positional argument)
        3. piped stdin
        4. sample_message.json  (so a bare `python main.py` still works)
    """
    # 1) Inline string.
    if args.message:
        return args.message

    # 2) A file path.
    file_path = args.file or args.path
    if file_path:
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"Message file not found: {path}")
        return path.read_text(encoding="utf-8")

    # 3) Piped stdin (isatty() is False when data is piped in).
    if not sys.stdin.isatty():
        piped = sys.stdin.read().strip()
        if piped:
            return piped

    # 4) Fall back to the bundled sample (sits next to main.py / project root).
    sample = Path(__file__).resolve().parent.parent / "sample_message.json"
    if sample.exists():
        logging.warning("No message provided; using sample_message.json for a test run.")
        return sample.read_text(encoding="utf-8")

    raise ValueError(
        "No message provided. Pass a file path, use --message, or pipe JSON via stdin."
    )


def run() -> int:
    """The local driver: parse args -> read message -> run pipeline -> report.

    Returns:
        An exit code (0 = success, 1 = failure) for `sys.exit` in main.py.
    """
    args = build_arg_parser().parse_args()
    configure_logging(args.verbose)

    try:
        message = read_message(args)
        result = process_message(message)  # the SAME shared pipeline
        logging.info("Done.")
        print("\nRESULT:")
        print(json.dumps(result, indent=2))
        return 0
    except Exception as exc:  # noqa: BLE001
        logging.error("Processing failed: %s", exc, exc_info=True)
        return 1
