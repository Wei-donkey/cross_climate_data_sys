# -*- coding: utf-8 -*-
"""Shared Oracle DB engine helper for the CROSS scripts.

Every script previously duplicated this block inline (read Config_DB.ini via
configparser, assemble an oracle:// connection string, create_engine). This
module centralizes it so a credential/driver change only needs updating here.
"""
from sqlalchemy import create_engine

from common.config import load_ini


def get_engine(section, ini_file='Config_DB.ini', echo=False):
    """Build a SQLAlchemy engine for an Oracle connection defined in Config_DB.ini."""
    section_cfg = load_ini(ini_file)[section]
    conn_string = (
        'oracle://' + section_cfg['user'] + ':' + section_cfg['password']
        + '@' + section_cfg['host'] + ':' + section_cfg['port'] + '/' + section_cfg['service']
    )
    return create_engine(conn_string, encoding='utf-8', echo=echo)
