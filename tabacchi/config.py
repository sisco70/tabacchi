#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# Copyright (C) Francesco Guarnieri 2020 <francesco@guarnie.net>
#

# import os
from pathlib import Path
import configparser
from importlib import metadata
import logging
import logging.config
import logging.handlers
from appdirs import AppDirs

# Estraggo il percorso principale da __file__ e deduco il nome del package principale
# (config.py deve stare nel package principale)
BASE_PATH = Path(__file__).resolve().parent

PACKAGE_NAME = BASE_PATH.name

dirs = AppDirs(appname=PACKAGE_NAME)

RESOURCE_PATH = BASE_PATH / 'resources'

user_data_dir = Path(dirs.user_data_dir)
user_config_dir = Path(dirs.user_config_dir)
user_log_dir = Path(dirs.user_log_dir)

# Se non esistono le directory standard per la configurazione, dati, logging
if not user_data_dir.exists():
    user_data_dir.mkdir(parents=True)
if not user_config_dir.exists():
    user_config_dir.mkdir(parents=True)
if not user_log_dir.exists():
    user_log_dir.mkdir(parents=True)

CONF_PATHNAME = user_config_dir / 'preferences.cfg'

__MINIMAL_LOGGING_CFG = f'''
[loggers]
keys=root,{PACKAGE_NAME}

[handlers]
keys=consoleHandler,fileHandler

[formatters]
keys=simpleFormatter

[logger_root]
level=WARNING
handlers=consoleHandler

[logger_{PACKAGE_NAME}]
level=WARNING
handlers=consoleHandler,fileHandler
qualname={PACKAGE_NAME}
propagate=0

[handler_consoleHandler]
class=StreamHandler
level=WARNING
formatter=simpleFormatter
args=(sys.stderr,)

[handler_fileHandler]
class=handlers.RotatingFileHandler
level=WARNING
formatter=simpleFormatter
args=("{user_log_dir / PACKAGE_NAME}.log", "a", 16384, 10)

[formatter_simpleFormatter]
format=%(asctime)s - %(name)s - %(levelname)s - %(message)s
datefmt=%Y-%m-%d %H:%M:%S
'''


def __initialize_logger():
    logging_file_config = user_config_dir / "logging.cfg"
    if not logging_file_config.exists():
        with open(logging_file_config, "w") as cfg_file:
            cfg_file.write(__MINIMAL_LOGGING_CFG)
    logging.config.fileConfig(logging_file_config)
    logger = logging.getLogger(PACKAGE_NAME)

    return logger


def __initialize_dunders():
    prj_filename = BASE_PATH.parent / "pyproject.toml"
    result = {'version': '0.0.0', 'author': '', 'desc': ''}
    if prj_filename.exists():
        config = configparser.ConfigParser()
        config.read(prj_filename)
        prj_data = config['tool.poetry']
        result['version'] = prj_data['version'].strip('"')
        result['author'] = prj_data['authors'].strip('[]').replace('"', '')
        result['desc'] = prj_data['description'].strip('"')
    else:
        try:
            program_metadata = metadata.metadata(PACKAGE_NAME)
            result['version'] = program_metadata["version"]
            result['author'] = program_metadata["author"]
            result['desc'] = program_metadata["summary"]
        except metadata.PackageNotFoundError:
            pass

    return result


dunders = __initialize_dunders()

__author__ = dunders['author']
__desc__ = dunders['desc']
__version__ = dunders['version']
__copyright__ = "Copyright 2020 - Francesco Guarnieri"

log = __initialize_logger()
