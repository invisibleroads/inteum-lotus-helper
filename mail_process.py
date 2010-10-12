'Classes for processing messages from mailbox'
# Import system modules
import os
import re
import shutil
import tempfile
import traceback
# Import custom modules
import log
import config
import query_process
import remark_process
import message_format
import model_database
import model_notes
import ticktock
import store


class Processor(object):

    # Constructor

    def __init__(self):
        # Connect
        self.mailbox = model_notes.Model(config.mailHost, config.mailPath, config.mailPassword)
        self.database = model_database.Model()
        # Make sure only one of us is running
        self.processID = setProcessID()
        # Make sure that necessary folders exist
        self.mailbox.enableFolder(config.folderFailed)
        self.mailbox.enableFolder(config.folderQuery)
        self.mailbox.enableFolder(config.folderRemark)
        self.mailbox.enableFolder(config.folderReservation)
        self.mailbox.enableFolder(config.folderUnauthorized)
        # Initialize
        self.document = None
        self.message = None
        self.temporaryFolders = []

    # Make

    def makeTemporaryFolder(self):
        temporaryFolder = tempfile.mkdtemp()
        self.temporaryFolders.append(temporaryFolder)
        return temporaryFolder

    # Process

    def process(self):
        # If a new process has started,
        if self.processID != getProcessID(): 
            # Quit
            raise ticktock.HaltError()
        # For each document,
        for self.document in self.mailbox.read():
            # Load fields
            self.message = model_notes.extractFields(self.document)
            # Screen email
            if self.processUnauthorized():
                continue
            # Try different approaches
            try:
                if self.processQuery(): 
                    continue
                if self.processRemark():
                    continue
                if self.processReservation():
                    continue
                if self.processRetrieval():
                    continue
                self.processSyntaxError()
            except:
                self.processFatalError()
        # Process queues
        self.mailbox.processQueues()
        # Delete temporary folders
        for temporaryFolder in self.temporaryFolders:
            if os.path.exists(temporaryFolder):
                shutil.rmtree(temporaryFolder)

    def processUnauthorized(self):
        'Process messages that do not come from an authorized address'
        # If the message is not from an authorized address,
        if self.message['from'] not in config.appUsers:
            # Move it
            self.queueMove(config.folderUnauthorized)
            # Return
            return True

    def processQuery(self): 
        'Process messages where the user is requesting information'
        return query_process.Processor(self).process()

    def processRemark(self):
        'Process messages where the user is adding a remark'
        return remark_process.Processor(self).process()

    def processReservation(self):
        'Process messages where the user is reserving a case number'
        # Test for reservation
        match = re.match(r'=\s*(.*)', self.message['subject'])
        # If we have a reservation,
        if match:
            # Extract title
            technologyTitle = match.group(1).strip()
            # Add technology
            valueByName = self.database.addTechnology(technologyTitle)
            # Send results
            self.queueWrite(message_format.templateReservation % valueByName)
            # Move to folderReservation
            self.queueMove(config.folderReservation)
            # Return
            return True

    def processRetrieval(self):
        'Process messages where the user is retrieving a document'
        # Test for retrieval
        match = re.match(r'doc\s*(.*)', self.message['subject'], re.I)
        # If we have a retrieval,
        if match:
            # Initialize
            errors = []
            summaries = []
            attachmentPaths = []
            temporaryFolder = self.makeTemporaryFolder()
            # For each documentID,
            for documentID in match.group(1).split():
                try:
                    documentID = int(documentID)
                # If the documentID is not an integer,
                except ValueError:
                    errors.append('Could not find document corresponding to documentID=%s' % documentID)
                    continue
                # Get documentContent
                documentContent = self.database.getDocumentContent(documentID)
                # If the document exists,
                if documentContent:
                    # Load
                    document, documentLink = self.database.getDocumentBundle(documentID)
                    # Format
                    summaries.append(message_format.formatDocumentBundle((document, documentLink)))
                    # Save
                    attachmentPath = os.path.join(temporaryFolder, '%s.%s' % (document['documentName'], document['fileExtension']))
                    attachmentPaths.append(attachmentPath)
                    with open(attachmentPath, 'wb') as attachmentFile:
                        attachmentFile.write(documentContent)
                # If the document does not exist,
                else:
                    errors.append('Could not find document corresponding to documentID=%s' % documentID)
                    continue
            # Send results
            self.queueWrite(message_format.templateRetrieval % {
                'errors': '<br>\n'.join(errors) + '<br><br>' if errors else '',
                'summaries': '<br><br>\n'.join(summaries),
            }, attachmentPaths)
            # Move to folderRetrieval
            self.queueMove(config.folderRetrieval)
            # Return
            return True

    def processSyntaxError(self):
        'Process messages that do not fit an expected category'
        self.queueWrite(message_format.formatErrorSyntax(self.message))
        self.queueMove(config.folderFailed)

    def processFatalError(self):
        'Process fatal errors'
        error = traceback.format_exc()
        log.error(error)
        self.queueWrite(message_format.formatErrorFatal(self.message, error))
        self.queueMove(config.folderFailed)

    # Queue

    def queueWrite(self, content, attachmentPaths=None):
        self.mailbox.queueWrite((self.message['from'], 'Re: ' + self.message['subject'], content, attachmentPaths))

    def queueMove(self, folder):
        self.mailbox.queueMove((folder, self.document))

    # Extract

    def extractAttachments(self):
        return model_notes.extractAttachments(self.document)


# ProcessID

def getProcessPath():
    return os.path.join(store.makeFolderSafely(config.appPath), 'process.id')

def getProcessID():
    # Get path
    processPath = getProcessPath()
    try:
        return int(open(processPath, 'rt').read())
    except:
        return 0

def setProcessID():
    # Get a new processID
    processID = getProcessID() + 1
    # Create a process file
    processFile = open(getProcessPath(), 'wt')
    processFile.write(str(processID))
    processFile.close()
    # Return
    return processID
