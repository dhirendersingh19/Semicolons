import logging

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

file_handler = logging.FileHandler('Logs/app.log')
file_handler.setLevel(logging.DEBUG)

logger.addHandler(file_handler)
