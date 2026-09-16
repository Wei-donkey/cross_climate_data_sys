# -*- coding: utf-8 -*-
"""Shared ini-file loading for the CROSS scripts."""
import configparser
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parent.parent


def load_ini(filename):
    """Read an ini file from the project root (utf-8-sig matches the existing files)."""
    config = configparser.ConfigParser()
    config.read(PROJECT_DIR / filename, encoding='utf-8-sig')
    return config
