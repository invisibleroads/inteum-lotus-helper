'Database, file and string storage'
# Import system modules
import os
import datetime
import cgi


# Database

def commit(function):
    def wrapper(*args, **kwargs):
        connection = executeWrite(function, args, kwargs)
        connection.commit()
    return wrapper

def executeRead(function, args, kwargs):
    # Set
    self = args[0]
    connection = self.connection1
    cursor = self.cursor1
    # Execute
    z = function(*args, **kwargs)
    sql, value = z[:2]
    if value: 
        cursor.execute(sql, value)
    else: 
        cursor.execute(sql)
    # Extract method
    method = z[2] if len(z) > 2 else None
    # Return
    return cursor, method

def executeWrite(function, args, kwargs):
    # Set
    self = args[0]
    connection = self.connection2
    cursor = self.cursor2
    # Execute
    sql, value = function(*args, **kwargs)
    cursor.execute(sql, value)
    return connection

def fetchOne(function):
    def wrapper(*args, **kwargs):
        cursor, method = executeRead(function, args, kwargs)
        result = cursor.fetchone()
        return method(result) if method else result
    return wrapper 

def fetchAll(function):
    def wrapper(*args, **kwargs):
        cursor, method = executeRead(function, args, kwargs)
        results = cursor.fetchall()
        return map(method, results) if method else results
    return wrapper 

def pullFirst(result):
    if result: 
        return result[0]


# File

basePath = os.path.dirname(os.path.abspath(__file__))

def loadTemplate(fileName):
    return open(os.path.join(basePath, 'templates', fileName), 'rt').read()

def makeFolderSafely(folderPath):
    'Make a directory at the given folderPath' 
    # For each parentPath, 
    for parentPath in reversed(traceParentPaths(os.path.abspath(folderPath))): 
        # If the parentPath folder does not exist, 
        if not os.path.exists(parentPath): 
            # Make the parentPath folder 
            os.mkdir(parentPath) 
    # Return 
    return folderPath 
 
def traceParentPaths(folderPath): 
    'Return a list of parentPaths containing the given folderPath' 
    parentPaths = [] 
    parentPath = folderPath 
    while parentPath not in parentPaths: 
        parentPaths.append(parentPath) 
        parentPath = os.path.dirname(parentPath) 
    return parentPaths 


# String

def strip(x):
    if hasattr(x, 'strip'):
        return x.strip()
    elif x == None:
        return ''
    else:
        return x

def escape(x):
    x = strip(x)
    if hasattr(x, 'strip'):
        return cgi.escape(x)
    else:
        return x


# When

def formatWhen(when, template):
    # If we have a string, parse a date from the string
    if type(when) == str:
        when = strip(when)
        if not when: 
            return ''
        when = datetime.datetime.strptime(when, '%m/%d/%Y')
    # Ignore invalid dates
    if not when or when.year < 1900: 
        return ''
    # Return
    return when.strftime(template)

def formatDate(when): 
    return formatWhen(when, '%b %d, %Y')

def formatDateTime(when): 
    return formatWhen(when, '%b %d, %Y %I:%M%p')
