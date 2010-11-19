'Report technologies, patents, agreements'
# Import system modules
import csv
import pyodbc
import datetime
import collections
# Import custom modules
import config
import imap
import mail
import store



def formatContactName(contactPack):
    # Expand
    (contactID, firstname, middleini, lastname, email), phonePacks = contactPack
    # Return
    return ((firstname.strip() + ' ' + middleini).strip() + ' ' + lastname).strip()


def formatPhone(phonePack):
    phoneID, phoneNumber, phoneType = phonePack
    return '(%s) %s' % (phoneType, phoneNumber)


def formatContactRow(contactPack):
    # Expand
    (contactID, firstname, middleini, lastname, email), phonePacks = contactPack
    firstname = firstname.strip()
    middleini = middleini.strip()
    lastname = lastname.strip()
    email = store.strip(email)
    # Format name
    parts = [] 
    if lastname:
        parts.append(lastname)
    if firstname:
        parts.append((firstname + ' ' + middleini).strip())
    name = ', '.join(parts)
    # Return
    return [name, email, ', '.join(formatPhone(x) for x in phonePacks)]


def formatTechnologyRow(technologyPack):
    # Expand
    (technologyID, technologyCaseNumber, technologyName, managerName, technologyStatus), interestPacks, departmentPacks, inventorIDs = technologyPack
    # Return
    return [
        technologyCaseNumber.strip(),
        technologyStatus.strip() if technologyStatus else '',
        managerName.strip(),
        ', '.join(x[0] for x in interestPacks).strip(),
        ', '.join(x[0] for x in departmentPacks).strip(),
        '',
        '',
        '',
        ', '.join(contactNameByID[x] for x in inventorIDs),
        technologyName.strip(),
    ]


def formatPatentRow(patentPack):
    # Expand
    (patentID, technologyCaseNumber, patentName, patentApplicationType, lawFirmName, lawFirmReferenceNumber, patentFilingDate, patentSerialNumber, patentNumber, patentStatus), inventorIDs = patentPack
    # Return
    return [
        store.strip(technologyCaseNumber),
        store.strip(patentStatus),
        store.strip(patentApplicationType),
        store.strip(lawFirmName),
        store.strip(lawFirmReferenceNumber),
        store.formatDate(patentFilingDate),
        store.strip(patentSerialNumber),
        store.strip(patentNumber),
        ', '.join(contactNameByID[x] for x in inventorIDs),
        store.strip(patentName),
    ]


# Connect to the database
connection = pyodbc.connect('dsn=' + config.databaseDSN)
cursor = connection.cursor()


# Load contacts
cursor.execute('SELECT primarykey, firstname, middleini, lastname, email FROM contacts ORDER BY lastname, firstname, middleini')
contactPacks = []
# For each contact,
for result in cursor.fetchall():
    # Load contactID
    contactID = result[0]
    # Load phone numbers
    cursor.execute('SELECT primarykey, phonenum, phonetype FROM phones WHERE contactsfk=?', [contactID])
    phonePacks = cursor.fetchall()
    # Append
    contactPacks.append([result, phonePacks])
# Build contactNameByID
contactNameByID = dict((x[0][0], formatContactName(x)) for x in contactPacks)


# Load technologies
cursor.execute('SELECT technol.primarykey, technol.techid, technol.name, users.loginid, techstat.name FROM technol LEFT JOIN users ON technol.managerfk=users.primarykey LEFT JOIN techstat ON technol.techstatfk=techstat.primarykey ORDER BY technol.techid')
technologyPacks = []
# For each technology,
for result in cursor.fetchall():
    # Load technologyID
    technologyID = result[0]
    # Load interests
    cursor.execute('SELECT DISTINCT interest.name FROM interest LEFT JOIN techints ON interest.primarykey=techints.interestfk WHERE techints.technolfk=?', [technologyID])
    interestPacks = cursor.fetchall()
    # Load client departments
    cursor.execute('SELECT DISTINCT clntdept.name FROM clntdept LEFT JOIN techclnt ON clntdept.primarykey=techclnt.clntdeptfk WHERE techclnt.technolfk=?', [technologyID])
    departmentPacks = cursor.fetchall()
    # Load inventors
    cursor.execute('SELECT DISTINCT contacts.primarykey FROM contacts LEFT JOIN techinv ON contacts.primarykey=techinv.contactsfk WHERE techinv.technolfk=?', [technologyID])
    inventorIDs = [x[0] for x in cursor.fetchall()]
    # Append
    technologyPacks.append([result, interestPacks, departmentPacks, inventorIDs])
# Build technologiesByInventor
technologiesByInventorID = collections.defaultdict(list)
# For each technologyPack,
for technologyPack in technologyPacks:
    # For each inventorID,
    for inventorID in technologyPack[3]:
        # Store
        technologiesByInventorID[inventorID].append(technologyPack)


# Load patents
cursor.execute('SELECT patents.primarykey, technol.techid, patents.name, papptype.name, company.name, patents.legalrefno, patents.filedate, patents.serialno, patents.patentno, patstat.name FROM patents LEFT JOIN company ON company.primarykey=patents.lawfirmfk LEFT JOIN papptype ON papptype.primarykey=patents.papptypefk LEFT JOIN technol ON technol.primarykey=patents.technolfk LEFT JOIN patstat ON patstat.primarykey=patents.patstatfk ORDER BY technol.techid, patents.filedate')
patentPacks = []
# For each patent,
for result in cursor.fetchall():
    # Load patentID
    patentID = result[0]
    # Load inventors
    cursor.execute('SELECT DISTINCT contacts.primarykey FROM contacts LEFT JOIN patinv ON contacts.primarykey=patinv.contactsfk LEFT JOIN techinv ON contacts.primarykey=techinv.contactsfk LEFT JOIN technol ON techinv.technolfk=technol.primarykey LEFT JOIN patents ON technol.primarykey=patents.technolfk WHERE patinv.patentsfk=? OR patents.primarykey=?', [patentID, patentID])
    inventorIDs = [x[0] for x in cursor.fetchall()]
    # Append
    patentPacks.append([result, inventorIDs])
# Build patentsByInventor
patentsByInventorID = collections.defaultdict(list)
# For each patentPack,
for patentPack in patentPacks:
    # For each inventorID,
    for inventorID in patentPack[1]:
        # Store
        patentsByInventorID[inventorID].append(patentPack)


# Load agreements
# Load marketing targets
# Load marketing projects


attachmentPaths = []
# Generate byInventor.csv
attachmentPaths.append('byInventor.csv')
csvWriter = csv.writer(open(attachmentPaths[-1], 'wb'))
# For each contactPack,
for contactPack in contactPacks:
    # Write contact
    csvWriter.writerow(formatContactRow(contactPack))
    # Get contactID
    contactID = contactPack[0][0]
    # If the contact has associated technologies,
    if contactID in technologiesByInventorID:
        # Write technologies
        csvWriter.writerow(['', 'Technologies'])
        csvWriter.writerow(['', '', 'Case #', 'Status', 'Manager', 'Interest', 'Campus', '', '', '', 'Inventors', 'Title'])
        for technologyPack in technologiesByInventorID[contactID]:
            csvWriter.writerow(['', ''] + formatTechnologyRow(technologyPack))
    # If the contact has associated patents,
    if contactID in patentsByInventorID:
        # Write patents
        csvWriter.writerow(['', 'Patents'])
        csvWriter.writerow(['', '', 'Case #', 'Status', 'Type', 'Law Firm', 'Reference #', 'Filing Date', 'Serial #', 'Patent #', 'Inventors', 'Title'])
        for patentPack in patentsByInventorID[contactID]:
            csvWriter.writerow(['', ''] + formatPatentRow(patentPack))


# Generate byLawFirm.csv
# Generate byCompany.csv


# For each reportRecipient,
for reportHost, reportEmail, reportUsername, reportPassword in config.reportRecipients:
    # Login to mail server
    imapStore = imap.Store(reportHost, reportUsername, reportPassword)
    # Send mail with attachments
    messageWhen = datetime.datetime.now()
    message = mail.buildMessage(
        {
            'From': config.mailAddress,
            'To': reportEmail,
            'Subject': 'Daily Report',
        }, 
        when=messageWhen,
        bodyText='Please see attachments.\n\n',
        attachmentPaths=attachmentPaths,
    )
    code, message = imapStore.revive('revived', message, messageWhen)
