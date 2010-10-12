'Simple framework for logging'
# Import system modules
import logging
import os
# Import custom modules
import config
import model_notes
import store


# Prepare
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s %(message)s', filename=os.path.join(store.makeFolderSafely(config.logPath), 'log.txt'), filemode='a')


def error(message):
    # Save in log
    logging.error(message)
    # Show on screen
    print message
    # Email administrator
    model_notes.Model(config.mailHost, config.mailPath, config.mailPassword).write(config.logEmail, '[tco-helper] Fatal error', message)
