import os
from logging import Logger, getLogger, FileHandler, StreamHandler
from logging.config import dictConfig

CONSOLE_HANDLER_NAME = 'console'
"""Default console handler name"""
FILE_HANDLER_NAME = 'file'
"""Default file handler name"""

config_logger = {
    "version": 1,
    "disable_existing_loggers": False,

    "formatters": {
        "simple": {
            "format": "%(asctime)s %(levelname)-8s : %(message)s",
            "datefmt": "%Y-%m-%d %H:%M:%S"
        },
        "verbose": {
            "format": "%(asctime)s %(name)-10s %(filename)-10s %(lineno)4d %(levelname)-8s : %(message)s",
            "datefmt": "%Y-%m-%d %H:%M:%S"
        },
    },

    "handlers": {
        CONSOLE_HANDLER_NAME: {
            "class": "logging.StreamHandler",
            "level": "INFO",
            "formatter": "simple",
            "stream": "ext://sys.stderr",
        },

        FILE_HANDLER_NAME: {
            "class": "logging.FileHandler",
            "level": "ERROR",
            "formatter": "verbose",
            # "filename": "./log/app.log",
            "encoding": "utf-8",
            "mode": "a",
        },

        # Example:
        # "file_size_rotate": {
        #     "class": "logging.handlers.RotatingFileHandler",
        #     "level": "WARNING",
        #     "formatter": "verbose",
        #     "filename": "./log/app.log",
        #     "encoding": "utf-8",
        #     "mode": "a",
        #     "maxBytes": 1024,
        #     "backupCount": 3,
        # },

        # Example:
        # "file_time_rotate": {
        #     "class": "logging.handlers.TimedRotatingFileHandler",
        #     "level": "WARNING",
        #     "formatter": "verbose",
        #     "filename": "./log/app.log",
        #     "encoding": "utf-8",
        #     "interval": 1,
        #     "backupCount": 3,
        # },
    },

    "root": {
        "level": "DEBUG",
        "handlers": [CONSOLE_HANDLER_NAME, FILE_HANDLER_NAME],
        # If more handlers are needed, add them on above list
    },
}
"""
Logging configuration sample dictionary
"""

def init_logger(logger_name: str, logfile_path: str|None=None) -> Logger:
    """
    Initialize and return a logger with the specified logger name and optional logging file path.

    Args:
        logger_name (str): Name of the logger to be created.
        logfile_path (str|None, optional): Path name to the log file. The directory will be created if it does not exist.

    Returns:
        Logger: Configured logger instance.

    Raises:
        Exception: If there is an error during logger configuration.
    """

    if logfile_path:
        directory = os.path.dirname(logfile_path)
        if not os.path.exists(directory):
            os.makedirs(directory)
        config_logger["handlers"][FILE_HANDLER_NAME]["filename"] = logfile_path

    try:
        dictConfig(config_logger)
        logger = getLogger(logger_name)
    except Exception as e:
        # If doing something is needed in the future, handle it here
        raise e

    return logger

def set_root_log_level(level: int):
    """
    Set the logging level for the root logger.

    Args:
        level (int): Logging level of root logger to set.
    """

    root_logger = getLogger()
    root_logger.setLevel(level)

def set_console_log_level(level: int):
    """
    Set the logging level for the console handler.

    Args:
        level (int): Logging level of console handler to set.
    """

    root_logger = getLogger()
    for handler in root_logger.handlers:
        if isinstance(handler, StreamHandler) and handler.get_name() == CONSOLE_HANDLER_NAME:
            handler.setLevel(level)
            break

def set_file_log_level(level: int):
    """
    Set the logging level for the file handler.

    Args:
        level (int): Logging level of file handler to set.
    """

    root_logger = getLogger()
    for handler in root_logger.handlers:
        if isinstance(handler, FileHandler) and handler.get_name() == FILE_HANDLER_NAME:
            handler.setLevel(level)
            break
