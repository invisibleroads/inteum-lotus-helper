'Command-line utility to run mail agent in a loop'
# Import system modules
import optparse
import traceback
# Import custom modules
import config
import ticktock
import mail_process
import log


if __name__ == '__main__':
    # Parse
    optionParser = optparse.OptionParser()
    optionParser.add_option('-1', '--once', dest='loop', action='store_false', default=True, help='Process once instead of continuously')
    options = optionParser.parse_args()[0]
    try: 
        # Connect
        processor = mail_process.Processor()
        # If we are running in a loop,
        if options.loop:
            # Loop
            loop = ticktock.Ticktock(config.appIntervalInSeconds, processor.process)
            loop.start()
        # If we are running once,
        else:
            # Process once
            processor.process()
    except: 
        # Record traceback
        log.error(traceback.format_exc())
