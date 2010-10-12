'Simple framework to run a function repeatedly at regular intervals'
# Import system modules
import threading
import time


class Ticktock(threading.Thread):
    'Class to run a function repeatedly at regular intervals'

    def __init__(self, intervalInSeconds, function):
        # Call super constructor
        super(Ticktock, self).__init__()
        # Set
        self.intervalInSeconds = intervalInSeconds
        self.function = function

    def run(self):
        'Run'
        while True:
            try: 
                self.function()
            except HaltError:
                break
            time.sleep(self.intervalInSeconds)


class HaltError(Exception):
    'Signal to halt'
    pass
