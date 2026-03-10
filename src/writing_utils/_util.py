"""
_util.py - Shared utilities for WritingUtils tools.
"""

import logging
import sys
import yaml


# ---------------------------------------------------------------------------
# Logging setup
# ---------------------------------------------------------------------------

_LOG_LEVELS = {
    "NONE":  None,
    "ERROR": logging.ERROR,
    "INFO":  logging.INFO,
    "DEBUG": logging.DEBUG,
}


def setup_logging(level_str, log_file=None):
    """Configure the root logger.  Called once after args are fully merged.

    level_str : one of NONE | ERROR | INFO | DEBUG
    log_file  : path to write log to; None → stderr
    """
    level = _LOG_LEVELS.get((level_str or "NONE").upper())
    if level is None:
        return  # logging disabled — leave root logger at default WARNING

    handler = (
        logging.FileHandler(log_file, encoding="utf-8")
        if log_file
        else logging.StreamHandler(sys.stderr)
    )
    handler.setFormatter(logging.Formatter(
        "%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    ))
    root = logging.getLogger()
    root.setLevel(level)
    root.addHandler(handler)


# ---------------------------------------------------------------------------
# Config file
# ---------------------------------------------------------------------------

def load_config(path):
    try:
        with open(path) as f:
            return yaml.safe_load(f) or {}
    except FileNotFoundError:
        print(f"Error: config file '{path}' not found", file=sys.stderr)
        sys.exit(1)
    except yaml.YAMLError as e:
        print(f"Error parsing config file: {e}", file=sys.stderr)
        sys.exit(1)
