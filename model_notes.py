'Methods for interacting with Lotus Notes'
# Import system modules
import win32com.client
import win32com.client.makepy
import pywintypes
import tempfile
import datetime
import os
# Import custom modules
import config
import store


# Generate type library
win32com.client.makepy.GenerateFromTypeLibSpec('Lotus Domino Objects')
win32com.client.makepy.GenerateFromTypeLibSpec('Lotus Notes Automation Classes')


# Interface for Lotus Notes
class Model(object):

    # Constructor

    def __init__(self, mailHost, mailName, mailPassword):
        # Connect
        self.notesSession = win32com.client.Dispatch('Lotus.NotesSession')
        # Try using a barrage of passwords
        try: 
            self.notesSession.Initialize(mailPassword)
            self.notesDatabase = self.notesSession.GetDatabase(mailHost, mailName)
        except pywintypes.com_error: 
            raise NotesError('Cannot access mail using %s on %s' % (mailName, mailHost))
        # Set constants
        self.ENC_NONE = win32com.client.constants.ENC_NONE
        self.EMBED_ATTACHMENT = win32com.client.constants.EMBED_ATTACHMENT
        # Initialize queues
        self.clearQueues()

    # Read

    def read(self, folder=None):
        # Load view
        view = self.notesDatabase.GetView(folder or config.folderInbox)
        # Get first document
        document = view.GetFirstDocument()
        # For each document,
        while document:
            # Yield
            yield document
            # Get the next document
            document = view.GetNextDocument(document)

    # Write

    def write(self, toWhom, subject, body, attachmentPaths=None):
        # Don't convert MIME to rich text
        self.notesSession.ConvertMime = False
        # Prepare message
        message = self.notesDatabase.CreateDocument()
        message.ReplaceItemValue('SendTo', toWhom)
        message.ReplaceItemValue('Subject', subject)
        stream = self.notesSession.CreateStream()
        stream.WriteText(body)
        messageBody = message.CreateMIMEEntity()
        messageBody.SetContentFromText(stream, 'text/html; charset=utf-8', self.ENC_NONE)
        # Add attachments
        for attachmentIndex, attachmentPath in enumerate(attachmentPaths or []):
            messageAttachment = message.CreateRichTextItem('attachment%d' % attachmentIndex)
            messageAttachment.EmbedObject(self.EMBED_ATTACHMENT, '', attachmentPath, None)
        # Send message
        message.Send(False)
        # Restore setting
        self.notesSession.ConvertMime = True

    # Move

    def move(self, destinationFolderName, document):
        # Move
        document.PutInFolder(destinationFolderName)
        document.RemoveFromFolder(config.folderInbox)

    # Queue

    def clearQueues(self):
        self.writes = []
        self.moves = []

    def queueWrite(self, parcel):
        self.writes.append(parcel)

    def queueMove(self, parcel):
        self.moves.append(parcel)

    def processQueues(self):
        for parcel in self.writes: 
            self.write(*parcel)
        for parcel in self.moves: 
            self.move(*parcel)
        self.clearQueues()

    # Enable

    def enableFolder(self, folderName):
        self.notesDatabase.EnableFolder(folderName)


def extractAttachments(document):
    'Extract attachment filenames and their contents'
    # Prepare
    attachmentPacks = []
    # For each item,
    for whichItem in xrange(len(document.Items)):
        # Get item
        item = document.Items[whichItem]
        # If the item is an attachment,
        if item.Name == '$FILE':
            # Prepare
            attachmentID, attachmentPath = tempfile.mkstemp()
            os.close(attachmentID)
            # Get the attachment
            fileName = item.Values[0]
            attachment = document.GetAttachment(fileName)
            attachment.ExtractFile(attachmentPath)
            attachmentContent = open(attachmentPath, 'rb').read()
            os.remove(attachmentPath)
            # Append
            attachmentPacks.append((fileName, attachmentContent))
    # Return
    return attachmentPacks


def extractFields(document):
    'Extract fields'
    # Prepare
    get = document.GetItemValue
    # Return
    return {
        'from': get('From')[0].strip(),
        'to': ', '.join(get('SendTo')),
        'when': datetime.datetime.fromtimestamp(int(get('PostedDate')[0])),
        'subject': get('Subject')[0].strip(),
        'body': get('Body')[0].strip(),
    }


class NotesError(Exception):
    pass
