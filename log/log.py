import logging
import os

loggers = {}

def setup_custom_logger(name):

    global loggers

    if loggers.get(name):
        return loggers.get(name)
    else:
        logger = logging.getLogger(name)

        fpath = os.path.dirname(os.path.abspath(__file__)) + r'\log_app_' + name + '.log'
        handler = logging.FileHandler(filename = fpath, mode='a',encoding=None,delay=False)

        formatter = logging.Formatter(fmt='%(asctime)s - %(levelname)s - %(module)s - %(message)s')
        handler.setFormatter(formatter)

        logger.addHandler(handler)
        logger.setLevel(logging.DEBUG)
        loggers[name] = logger

        return logger