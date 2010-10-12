'Configuration module'
# Import system modules
import ConfigParser


# Load default configuration file
default = ConfigParser.ConfigParser()
default.read('default.cfg')
# Load custom configuration file
custom = ConfigParser.ConfigParser()
custom.read('.default.cfg')


# Define method
def get(section, option, convert=str):
    'Get option'
    # Get
    configuration = custom if custom.has_option(section, option) else default
    # Return
    return convert(configuration.get(section, option))


# Load values
appIntervalInSeconds = get('app', 'interval in seconds', int)
appPath              = get('app', 'path')
appUsers             = get('app', 'users', lambda x: x.strip().splitlines())
databaseDSN          = get('database', 'dsn')
folderInbox          = get('folder', 'inbox')
folderFailed         = get('folder', 'failed')
folderQuery          = get('folder', 'query')
folderRemark         = get('folder', 'remark')
folderReservation    = get('folder', 'reservation')
folderRetrieval      = get('folder', 'retrieval')
folderTrash          = get('folder', 'trash')
folderUnauthorized   = get('folder', 'unauthorized')
logPath              = get('log', 'path')
logEmail             = get('log', 'email')
mailAddress          = get('mail', 'address')
mailHost             = get('mail', 'host')
mailPassword         = get('mail', 'password')
mailPath             = get('mail', 'path')
reportRecipients     = get('report', 'recipients', lambda x: [row.split() for row in x.strip().splitlines()])
