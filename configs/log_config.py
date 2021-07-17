import logging
from logging.handlers import TimedRotatingFileHandler
from pathlib import Path

# __current_config = None
__logger = None
filename = "pptfinder"


class BaseConfig:
    def __init__(self):
        # Load variables from environment
        self.LOG_FORMAT = '[%(asctime)s] [%(process)d] [%(levelname)s] [%(module)s.%(lineno)d.%(funcName)s] [%(threadName)s] %(message)s'
        self.LOG_FILE = Path(Path.cwd(), "log", f"{filename}.log")
        self.LOG_FILE_SUFFIX = '%Y-%m-%d'
        self.LOG_FILE_COUNT = 30
        self.LOG_LEVEL = 'DEBUG'
        self.LOG_WHEN = 'midnight'
        self.LOG_INTERVAL = 1

        # Set logger
        path = self.LOG_FILE.parent
        if not (path.exists() and path.is_dir()):
            Path.mkdir(path)
        logging.basicConfig(
            level=getattr(logging, self.LOG_LEVEL),
            format=self.LOG_FORMAT
        )
        self.logger = logging.getLogger()
        self.logger.addHandler(self.get_log_handler())

    def get_log_handler(self):
        file_handler = TimedRotatingFileHandler(self.LOG_FILE, when=self.LOG_WHEN,
                                                interval=self.LOG_INTERVAL,
                                                encoding='UTF-8', backupCount=self.LOG_FILE_COUNT)
        file_handler.suffix = self.LOG_FILE_SUFFIX
        file_formatter = logging.Formatter(self.LOG_FORMAT)
        file_handler.setFormatter(file_formatter)
        file_handler.level = getattr(logging, self.LOG_LEVEL)
        return file_handler


def get_logger():
    return __logger


def init_config(config_name=None) -> BaseConfig:
    global __logger
    __logger = BaseConfig().logger
