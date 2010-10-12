'Database routines'
# Import system modules
import os
import odbc
import pyodbc
import datetime
# Import custom modules
import store
import config
import message_format


class Model(object):
    'Class containing database routines'

    # Constructor

    def __init__(self):
        'Connect to the database'
        dsn = config.databaseDSN
        self.connection1 = pyodbc.connect('dsn=' + dsn)
        self.connection2 = odbc.odbc('dsn=' + dsn)
        self.cursor1 = self.connection1.cursor()
        self.cursor2 = self.connection2.cursor()
        self.getBundleByLabel = {
            'Agreements': self.getAgreementBundle,
            'Companies': self.getCompanyBundle,
            'Contacts': self.getContactBundle,
            'Documents': self.getDocumentBundle,
            'Marketing Projects': self.getMarketingProjectBundle,
            'Marketing Targets': self.getMarketingTargetBundle,
            'Patents': self.getPatentBundle,
            'Remarks': self.getRemarkBundle,
            'Technologies': self.getTechnologyBundle,
        }

    # Destructor

    def __del__(self):
        'Disconnect from the database'
        self.connection1.close()
        self.connection2.close()

    # General

    def add(self, table, fields, valueByName):
        'Add a record to a table'
        # Initialize
        errors = []
        now = makeNow()
        fields = [
            ('primarykey', self.getNextPrimaryKey(table), False),
            ('rcreatedd', now, False),
            ('rcreatedu', 1044, False),
            ('rupdatedd', now, False),
            ('rupdatedu', 1044, False),
        ] + fields
        # Fill valueByName
        for name, default, isRequired in fields:
            # If the value is not defined,
            if name not in valueByName:
                # If the value is required, append error
                if isRequired: 
                    errors.append(name + ' is required')
                # Set the default value
                valueByName[name] = default
        # Execute
        self.cursor2.execute('INSERT INTO %s (%s) VALUES (%s)' % (table, ','.join(x[0] for x in fields), ','.join(['?'] * len(fields))), [valueByName[x[0]] for x in fields])
        # If there are errors, raise an exception
        if errors: 
            raise ExtractionError('\n'.join(errors))

    def findIDs(self, table, joins, wheres, patterns, term, orderBys):
        'Find matching ids'
        # Build sql
        sql = 'SELECT DISTINCT %(fields)s FROM %(table)s %(joins)s WHERE %(wheres)s ORDER BY %(orderBys)s' % {
            'fields': ','.join([table + '.primarykey'] + [x.split()[0] for x in orderBys]),
            'table': table,
            'joins': ' '.join(joins),
            'wheres': ' OR '.join(list(wheres) + patterns),
            'orderBys': ','.join(orderBys),
        }
        # Execute
        self.cursor1.execute(sql, ['%' + term + '%'] * len(patterns))
        # Return
        return [x[0] for x in self.cursor1.fetchall()]

    def getLink(self, linkTable, linkID):
        if linkTable == 'agrmnts': 
            self.cursor1.execute(buildAgreementSQL(where='agrmnts.primarykey=?'), [linkID]) 
            return stripAgreement(self.cursor1.fetchone()) 
        if linkTable == 'company': 
            self.cursor1.execute(buildCompanySQL(where='company.primarykey=?'), [linkID]) 
            return stripCompany(self.cursor1.fetchone()) 
        if linkTable == 'contacts': 
            self.cursor1.execute(buildContactSQL(where='contacts.primarykey=?'), [linkID]) 
            return stripContact(self.cursor1.fetchone()) 
        if linkTable == 'mktprj': 
            self.cursor1.execute(buildMarketingProjectSQL(where='mktprj.primarykey=?'), [linkID]) 
            return stripMarketingProject(self.cursor1.fetchone()) 
        if linkTable == 'mkttgt': 
            self.cursor1.execute(buildMarketingTargetSQL(where='mkttgt.primarykey=?'), [linkID]) 
            return stripMarketingTarget(self.cursor1.fetchone()) 
        if linkTable == 'patents': 
            self.cursor1.execute(buildPatentSQL(where='patents.primarykey=?'), [linkID]) 
            return stripPatent(self.cursor1.fetchone()) 
        if linkTable == 'payable': 
            self.cursor1.execute(buildPayableSQL(where='payable.primarykey=?'), [linkID]) 
            return stripPayable(self.cursor1.fetchone()) 
        if linkTable == 'payment': 
            self.cursor1.execute(buildPaymentSQL(where='payment.primarykey=?'), [linkID]) 
            return stripPayment(self.cursor1.fetchone()) 
        if linkTable == 'recvable': 
            self.cursor1.execute(buildReceivableSQL(where='recvable.primarykey=?'), [linkID]) 
            return stripReceivable(self.cursor1.fetchone()) 
        if linkTable == 'remarks': 
            self.cursor1.execute(buildRemarkSQL(where='remarks.primarykey=?'), [linkID]) 
            return stripRemark(self.cursor1.fetchone()) 
        if linkTable == 'technol': 
            self.cursor1.execute(buildTechnologySQL(where='technol.primarykey=?'), [linkID]) 
            return stripTechnology(self.cursor1.fetchone()) 

    def getNextPrimaryKey(self, table):
        'Return next available primary key'
        # Select largest primarykey
        self.cursor1.execute('SELECT MAX(primarykey) FROM AUDITLOG')
        key = self.cursor1.fetchone()[0]
        result = True
        while result:
            key = key + 1
            self.cursor1.execute('SELECT primarykey FROM %s WHERE primarykey=?' % table, [key])
            result = self.cursor1.fetchone()
        return key

    def getNextRemarkOrder(self, linkTable, linkFK):
        self.cursor1.execute('SELECT MAX(orderno) FROM remarks WHERE linkfk=? AND linktable=?', (linkFK, linkTable))
        lastOrder = self.cursor1.fetchone()[0]
        return lastOrder + 1 if lastOrder else 1

    # Agreement

    def findAgreementIDs(self, term, companyIDString=None, contactIDString=None, technologyIDString=None):
        'Return matching agreements'
        # Initialize
        joins, wheres = set(), set()
        # Use companies
        if companyIDString == None:
            companyIDString = makeIDString(self.findCompanyIDs(term))
        if companyIDString:
            # As parties
            joins.add('LEFT JOIN agrparty ON agrmnts.primarykey=agrparty.agrmntsfk')
            wheres.add("(agrparty.linktable='company' AND agrparty.linkfk IN (%s))" % companyIDString)
        # Use contacts
        if contactIDString == None:
            contactIDString = makeIDString(self.findContactIDs(term))
        if contactIDString:
            # As researchers
            joins.add('LEFT JOIN agrrchr ON agrmnts.primarykey=agrrchr.agrmntsfk')
            wheres.add('agrrchr.contactsfk IN (%s)' % contactIDString)
            # As parties
            joins.add('LEFT JOIN agrparty ON agrmnts.primarykey=agrparty.agrmntsfk')
            wheres.add("(agrparty.linktable='contacts' AND agrparty.linkfk IN (%s))" % contactIDString)
        # Use technologies
        if technologyIDString == None:
            technologyIDString = makeIDString(self.findTechnologyIDs(term, contactIDString))
        if technologyIDString:
            # As technologies
            joins.add('LEFT JOIN agrtech ON agrmnts.primarykey=agrtech.agrmntsfk')
            wheres.add('agrtech.technolfk IN (%s)' % technologyIDString)
        # Build patterns
        patterns = [
            'agrmnts.agrmntid LIKE ?', 
            'agrmnts.name LIKE ?',
        ]
        # Return
        return self.findIDs('agrmnts', joins, wheres, patterns, term, ['agrmnts.agrmntid'])

    def getAgreementBundle(self, agreementID, simple=False):
        'Get information related to the agreement'
        # Set value
        value = [agreementID]
        # Get agreement
        self.cursor1.execute(buildAgreementSQL(where='agrmnts.primarykey=?'), value)
        agreement = stripAgreement(self.cursor1.fetchone())
        if simple:
            return agreement
        # Get agreementResearcherContactBundles
        self.cursor1.execute('SELECT contacts.primarykey FROM contacts INNER JOIN agrrchr ON contacts.primarykey=agrrchr.contactsfk WHERE agrrchr.agrmntsfk=?', value)
        agreementResearcherContactBundles = [self.getContactBundle(x[0]) for x in self.cursor1.fetchall()]
        # Get agreementPartyContactBundles
        self.cursor1.execute("SELECT linkfk FROM agrparty WHERE agrmntsfk=? AND linktable='contacts'", value)
        agreementPartyContactBundles = [self.getContactBundle(x[0]) for x in self.cursor1.fetchall()]
        # Get agreementPartyCompanyBundles
        self.cursor1.execute("SELECT linkfk FROM agrparty WHERE agrmntsfk=? AND linktable='company'", value)
        agreementPartyCompanyBundles = [self.getCompanyBundle(x[0]) for x in self.cursor1.fetchall()]
        # Get agreementRemarks
        agreementRemarks = self.getRemarks('agrmnts', agreementID)
        # Get agreementTechnologies
        self.cursor1.execute(buildTechnologySQL(join='INNER JOIN agrtech ON technol.primarykey=agrtech.technolfk', where='agrtech.agrmntsfk=?'), value)
        agreementTechnologies = [stripTechnology(x) for x in self.cursor1.fetchall()]
        # Get agreementPatents
        self.cursor1.execute(buildPatentSQL(join='INNER JOIN agrpat ON patents.primarykey=agrpat.patentsfk', where='agrpat.agrmntsfk=?'), value)
        agreementPatents = [stripPatent(x) for x in self.cursor1.fetchall()]
        # Return
        return agreement, agreementResearcherContactBundles, agreementPartyContactBundles, agreementPartyCompanyBundles, agreementRemarks, agreementTechnologies, agreementPatents

    # Audit

    def addAudit(self, linkTable, linkFK, identifier, details, actionType='A'):
        'Add a record to the table that keeps track of changes to the database'
        return self.addAuditFromDictionary({
            'linktable': linkTable,
            'linkfk': linkFK,
            'identifier': identifier[:50],
            'details': details,
            'type': actionType,
        })

    def addAuditFromDictionary(self, valueByName):
        'Add a record to the table that keeps track of changes to the database'
        # Add
        self.add('auditlog', [
            ('linktable', '', True),
            ('linkfk', 0, True),
            ('identifier', '', True),
            ('details', '', True),
            ('type', '', True),
        ], valueByName)
        # Return
        return valueByName

    # Company

    def findCompanyIDs(self, term):
        'Return matching companyIDs'
        # Build patterns
        patterns = [
            'name LIKE ?',
        ]
        # Return
        return self.findIDs('company', [], [], patterns, term, ['company.name'])

    def getCompanyBundle(self, companyID, simple=False):
        'Get information related to the company'
        # Set value
        value = [companyID]
        # Get company
        self.cursor1.execute(buildCompanySQL(where='company.primarykey=?'), value)
        company = stripCompany(self.cursor1.fetchone())
        if simple:
            return company
        # Get companyContactBundles
        self.cursor1.execute('SELECT primarykey FROM contacts WHERE companyfk=?', value)
        companyContactBundles = [self.getContactBundle(x[0]) for x in self.cursor1.fetchall()]
        # Get companyRemarks
        companyRemarks = self.getRemarks('company', companyID)
        # Return
        return company, companyContactBundles, companyRemarks

    # Contact

    def findContactIDs(self, term):
        'Return matching contactIDs'
        # Build patterns
        patterns = [
            "RTRIM(firstname + ' ' + middleini) + ' ' + lastname LIKE ?", 
            'email LIKE ?',
        ]
        # Return
        return self.findIDs('contacts', [], [], patterns, term, ['contacts.lastname', 'contacts.firstname', 'contacts.middleini'])

    def getContactBundle(self, contactID, simple=False):
        'Get information related to the contact'
        # Set value
        value = [contactID]
        # Get contact
        self.cursor1.execute(buildContactSQL(where='contacts.primarykey=?'), value)
        contact = stripContact(self.cursor1.fetchone())
        if simple:
            return contact
        # Get contactPhones
        self.cursor1.execute('SELECT primarykey, phonenum, phonetype FROM phones WHERE contactsfk=?', value)
        contactPhones = map(stripPhone, self.cursor1.fetchall())
        # Get contactRemarks
        contactRemarks = self.getRemarks('contacts', contactID)
        # Return
        return contact, contactPhones, contactRemarks
    
    # Document

    def addDocument(self, linkTable, linkFK, fileName, fileContent):
        'Add a document'
        # Split
        fileBase, fileExtension = os.path.splitext(fileName)
        # Return
        return self.addDocumentFromDictionary({
            'linktable': linkTable,
            'linkfk': linkFK,
            'location': '*' + fileName[:250],
            'name': fileBase[:250],
            'fileexten': fileExtension.lower().replace('.', '')[:5],
            'contents': fileContent,
        })

    def addDocumentFromDictionary(self, valueByName):
        'Add a document'
        # Add
        self.add('document', [
            ('name', '', True),
            ('linktable', '', True),
            ('linkfk', 0, True),
            ('location', '', True),
            ('compressed', 'N', False),
            ('keywords', '', False),
            ('origdelete', 'Y', False),
            ('crntcsum', 0, False),
            ('lstuplcsum', 0, False),
            ('doccatfk', None, False),
            ('fileexten', '', False),
            ('ownerfK', 1044, False),
            ('private', 'N', False),
            ('datapart', None, False),
            ('archive', None, False),
        ], valueByName)
        # Add audit
        self.addAudit('document', valueByName['primarykey'], 'Title:%s' % valueByName['name'], '[Name] [LinkFK] [OwnerFK] [Location] [Keywords] [LinkTable] [File Exten] [DocCatFK]')
        # Commit
        self.connection2.commit()
        # If we need to store contents,
        if 'contents' in valueByName:
            # Store contents via cursor1 to avoid strange error
            self.cursor1.execute('UPDATE document SET contents=? WHERE document.primarykey=?', (buffer(valueByName['contents']), valueByName['primarykey']))
            # Commit
            self.connection1.commit()
        # Return
        return valueByName

    def findDocumentIDs(self, term, agreementIDString=None, companyIDString=None, contactIDString=None, marketingProjectIDString=None, marketingTargetIDString=None, patentIDString=None, remarkIDString=None, technologyIDString=None):
        'Return matching documentIDs'
        # Initialize
        joins, wheres = set(), set()
        # Use agreements
        if agreementIDString == None:
            agreementIDString = makeIDString(self.findAgreementIDs(term))
        if agreementIDString:
            # As links
            wheres.add("(document.linktable='agrmnts' AND document.linkfk IN (%s))" % agreementIDString)
            # As parties
            joins.add("LEFT JOIN recvable ON document.linktable='recvable' AND document.linkfk=recvable.primarykey")
            wheres.add('recvable.agrpartyfk IN (%s)' % agreementIDString)
        # Use companies
        if companyIDString == None:
            companyIDString = makeIDString(self.findCompanyIDs(term))
        if companyIDString:
            # As links
            wheres.add("(document.linktable='company' AND document.linkfk IN (%s))" % companyIDString)
            # As payables
            joins.add("LEFT JOIN payable ON document.linktable='payable' AND document.linkfk=payable.primarykey")
            wheres.add("(payable.linktable='company' AND payable.linkfk IN (%s))" % companyIDString)
            # As payments
            joins.add("LEFT JOIN payment ON document.linktable='payment' AND document.linkfk=payment.primarykey")
            wheres.add("(payment.linktable='company' AND payment.linkfk IN (%s))" % companyIDString)
        # Use contacts
        if contactIDString == None:
            contactIDString = makeIDString(self.findContactIDs(term))
        if contactIDString:
            # As links
            wheres.add("(document.linktable='contacts' AND document.linkfk IN (%s))" % contactIDString)
            # As payables
            joins.add("LEFT JOIN payable ON document.linktable='payable' AND document.linkfk=payable.primarykey")
            wheres.add("(payable.linktable='contacts' AND payable.linkfk IN (%s))" % contactIDString)
            # As payments
            joins.add("LEFT JOIN payment ON document.linktable='payment' AND document.linkfk=payment.primarykey")
            wheres.add("(payment.linktable='contacts' AND payment.linkfk IN (%s))" % contactIDString)
        # Use marketing projects
        if marketingProjectIDString == None:
            marketingProjectIDString = makeIDString(self.findMarketingProjectIDs(term))
        if marketingProjectIDString:
            wheres.add("(document.linktable='mktprj' AND document.linkfk IN (%s))" % marketingProjectIDString)
        # Use marketing targets
        if marketingTargetIDString == None:
            marketingTargetIDString = makeIDString(self.findMarketingTargetIDs(term))
        if marketingTargetIDString:
            wheres.add("(document.linktable='mkttgt' AND document.linkfk IN (%s))" % marketingTargetIDString)
        # Use patents
        if patentIDString == None:
            patentIDString = makeIDString(self.findPatentIDs(term))
        if patentIDString:
            wheres.add("(document.linktable='patents' AND document.linkfk IN (%s))" % patentIDString)
        # Use remarks
        if remarkIDString == None:
            remarkIDString = makeIDString(self.findRemarkIDs(term))
        if remarkIDString:
            wheres.add("(document.linktable='remarks' AND document.linkfk IN (%s))" % remarkIDString)
        # Use technologies
        if technologyIDString == None:
            technologyIDString = makeIDString(self.findTechnologyIDs(term))
        if technologyIDString:
            wheres.add("(document.linktable='technol' AND document.linkfk IN (%s))" % technologyIDString)
        # Build patterns
        patterns = [
            'document.name LIKE ?',
        ]
        # Return
        return self.findIDs('document', joins, wheres, patterns, term, ['document.name'])

    def getDocumentBundle(self, documentID, simple=False):
        'Get information related to the document'
        # Set value 
        value = [documentID] 
        # Get document 
        self.cursor1.execute(buildDocumentSQL(where='document.primarykey=?'), value) 
        document = stripDocument(self.cursor1.fetchone()) 
        if simple:
            return document
        # Get documentLink 
        documentLink = self.getLink(document['linkTable'], document['linkID'])
        # Return 
        return document, documentLink 

    @store.fetchOne
    def getDocumentContent(self, documentID):
        'Retrieve document content'
        return 'SELECT contents FROM document WHERE primarykey=?', [documentID], store.pullFirst

    # Marketing project

    def findMarketingProjectIDs(self, term, companyIDString=None, contactIDString=None, patentIDString=None, technologyIDString=None):
        'Return matching marketing projects'
        # Initialize
        joins, wheres = set(), set()
        # Use contacts
        if contactIDString == None:
            contactIDString = makeIDString(self.findContactIDs(term))
        if contactIDString:
            joins.add('LEFT JOIN mkprjinv ON mktprj.primarykey=mkprjinv.mktprjfk')
            wheres.add('mkprjinv.contactsfk IN (%s)' % contactIDString)
        # Use technologies
        if technologyIDString == None:
            technologyIDString = makeIDString(self.findTechnologyIDs(term, contactIDString))
        if technologyIDString:
            joins.add('LEFT JOIN mkprjtec ON mktprj.primarykey=mkprjtec.mktprjfk')
            wheres.add('mkprjtec.technolfk IN (%s)' % technologyIDString)
        # Use patents
        if patentIDString == None:
            patentIDString = makeIDString(self.findPatentIDs(term, companyIDString, contactIDString, technologyIDString))
        if patentIDString:
            joins.add('LEFT JOIN mkprjpat ON mktprj.primarykey=mkprjpat.mktprjfk')
            wheres.add('mkprjpat.patentsfk IN (%s)' % patentIDString)
        # Build patterns
        patterns = [
            'mktprjid LIKE ?',
            'mktprj.name LIKE ?',
        ]
        # Return
        return self.findIDs('mktprj', joins, wheres, patterns, term, ['mktprj.name'])

    def getMarketingProjectBundle(self, marketingProjectID, simple=False):
        'Get information related to the marketing project'
        # Set value
        value = [marketingProjectID]
        # Get marketingProject
        self.cursor1.execute(buildMarketingProjectSQL(where='mktprj.primarykey=?'), value)
        marketingProject = stripMarketingProject(self.cursor1.fetchone())
        if simple:
            return marketingProject
        # Get marketingProjectMarketingTargetBundles
        self.cursor1.execute("SELECT mkttgt.primarykey FROM mkttgt WHERE linktable='mktprj' AND linkfk=?", value)
        marketingProjectMarketingTargetBundles = [self.getMarketingTargetBundle(x[0]) for x in self.cursor1.fetchall()]
        # Get marketingProjectContactBundles
        self.cursor1.execute('SELECT contacts.primarykey FROM contacts INNER JOIN mkprjinv ON contacts.primarykey=mkprjinv.contactsfk WHERE mkprjinv.mktprjfk=? ORDER BY pi_order', value)
        marketingProjectContactBundles = [self.getContactBundle(x[0]) for x in self.cursor1.fetchall()]
        # Get marketingProjectRemarks
        marketingProjectRemarks = self.getRemarks('mktprj', marketingProjectID)
        # Get marketingProjectTechnologies
        self.cursor1.execute(buildTechnologySQL(join='INNER JOIN mkprjtec ON technol.primarykey=mkprjtec.technolfk', where='mkprjtec.mktprjfk=?'), value)
        marketingProjectTechnologies = [stripTechnology(x) for x in self.cursor1.fetchall()]
        # Get marketingProjectPatents
        self.cursor1.execute(buildPatentSQL(join='INNER JOIN mkprjpat ON patents.primarykey=mkprjpat.patentsfk', where='mkprjpat.mktprjfk=?'), value)
        marketingProjectPatents = [stripPatent(x) for x in self.cursor1.fetchall()]
        # Return
        return marketingProject, marketingProjectMarketingTargetBundles, marketingProjectContactBundles, marketingProjectRemarks, marketingProjectTechnologies, marketingProjectPatents

    # Marketing target

    def findMarketingTargetIDs(self, term, companyIDString=None, contactIDString=None, marketingProjectIDString=None, patentIDString=None, technologyIDString=None):
        'Return matching marketing targets'
        # Initialize
        wheres = set()
        # Use companies
        if companyIDString == None:
            companyIDString = makeIDString(self.findCompanyIDs(term))
        if companyIDString:
            wheres.add('mkttgt.companyfk IN (%s)' % companyIDString)
        # Use contacts
        if contactIDString == None:
            contactIDString = makeIDString(self.findContactIDs(term))
        if contactIDString:
            wheres.add('mkttgt.contactsfk IN (%s)' % contactIDString)
        # Use technologies
        if technologyIDString == None:
            technologyIDString = makeIDString(self.findTechnologyIDs(term, contactIDString))
        if technologyIDString:
            wheres.add("(mkttgt.linktable='technol' AND mkttgt.linkfk IN (%s))" % technologyIDString)
        # Use patents
        if patentIDString == None:
            patentIDString = makeIDString(self.findPatentIDs(term, companyIDString, contactIDString, technologyIDString))
        if patentIDString:
            wheres.add("(mkttgt.linktable='patents' AND mkttgt.linkfk IN (%s))" % patentIDString)
        # Use marketing projects
        if marketingProjectIDString == None:
            marketingProjectIDString = makeIDString(self.findMarketingProjectIDs(term, companyIDString, contactIDString, patentIDString, technologyIDString))
        if marketingProjectIDString:
            wheres.add("(mkttgt.linktable='mktprj' AND mkttgt.linkfk IN (%s))" % marketingProjectIDString)
        # Return
        return self.findIDs('mkttgt', ['LEFT JOIN company ON company.primarykey=mkttgt.companyfk'], wheres, [], term, ['company.name']) if wheres else []

    def getMarketingTargetBundle(self, marketingTargetID, simple=False):
        'Get information related to the marketing target'
        # Set value
        value = [marketingTargetID]
        # Get marketingTarget
        self.cursor1.execute(buildMarketingTargetSQL(where='mkttgt.primarykey=?'), value)
        marketingTarget = stripMarketingTarget(self.cursor1.fetchone())
        if simple:
            return marketingTarget
        # Get marketingTargetLink
        marketingTargetLink = self.getLink(marketingTarget['linkTable'], marketingTarget['linkID'])
        # Get marketingTargetRemarks
        marketingTargetRemarks = self.getRemarks('mkttgt', marketingTargetID)
        # Get marketingTargetAgreements
        self.cursor1.execute(buildAgreementSQL(join='INNER JOIN mktgtagr ON agrmnts.primarykey=mktgtagr.agrmntsfk', where='mktgtagr.mkttgtfk=?'), value)
        marketingTargetAgreements = [stripAgreement(x) for x in self.cursor1.fetchall()]
        # Return
        return marketingTarget, marketingTargetLink, marketingTargetRemarks, marketingTargetAgreements

    # Patent

    def findPatentIDs(self, term, companyIDString=None, contactIDString=None, technologyIDString=None):
        'Return matching patentIDs'
        # Initialize
        joins, wheres = set(), set()
        # Use companies
        if companyIDString == None:
            companyIDString = makeIDString(self.findCompanyIDs(term))
        if companyIDString:
            wheres.add('patents.lawfirmfk IN (%s)' % companyIDString)
        # Use contacts
        if contactIDString == None:
            contactIDString = makeIDString(self.findContactIDs(term))
        if contactIDString:
            # As patent inventors
            joins.add('LEFT JOIN patinv ON patents.primarykey=patinv.patentsfk')
            wheres.add('patinv.contactsfk IN (%s)' % contactIDString)
            # As technology inventors
            joins.add('LEFT JOIN technol ON patents.technolfk=technol.primarykey')
            joins.add('LEFT JOIN techinv ON technol.primarykey=techinv.technolfk')
            wheres.add('techinv.contactsfk IN (%s)' % contactIDString)
        # Use technologies
        if technologyIDString == None:
            technologyIDString = makeIDString(self.findTechnologyIDs(term, contactIDString))
        if technologyIDString:
            wheres.add('patents.technolfk IN (%s)' % technologyIDString)
        # Build patterns
        patterns = [
            'patents.name LIKE ?', 
            'patents.legalrefno LIKE ?', 
            "REPLACE(patents.serialno, ',', '') LIKE ?", 
            "REPLACE(patents.patentno, ',', '') LIKE ?",
        ]
        # Return
        return self.findIDs('patents', joins, wheres, patterns, term, ['patents.filedate'])

    def getPatentBundle(self, patentID, simple=False):
        'Get information related to the patent'
        # Set value
        value = [patentID]
        # Get patent
        self.cursor1.execute(buildPatentSQL(where='patents.primarykey=?'), value)
        patent = stripPatent(self.cursor1.fetchone())
        if simple:
            return patent
        # Get patentContactBundles
        self.cursor1.execute('SELECT contacts.primarykey FROM contacts INNER JOIN patinv ON contacts.primarykey=patinv.contactsfk WHERE patinv.patentsfk=? ORDER BY pi_order', value)
        patentContactBundles = [self.getContactBundle(x[0]) for x in self.cursor1.fetchall()]
        patent['leadInventor'] = message_format.formatContactName(patentContactBundles[0][0]) if patentContactBundles else ''
        # Get patentRemarks
        patentRemarks = self.getRemarks('patents', patentID)
        # Get patentPayableDetails
        patentPayableDetails = self.getPayableDetailsForPatent(patentID)
        # Return
        return patent, patentContactBundles, patentRemarks, patentPayableDetails

    # Payable

    @store.fetchAll
    def getPayableDetailsForTechnology(self, technologyID):
        'Get payableDetails related to technology'
        sql = ' UNION '.join([
            buildPayableDetailSQL(where='paybldtl.linktable=? AND paybldtl.linkfk=?'),
            buildPayableDetailSQL(where='paybldtl.linktable=? AND patents.technolfk=?', join='LEFT JOIN patents ON patents.primarykey=paybldtl.linkfk'),
        ]) + ' ORDER BY payable.duedate DESC'
        return sql, ('technol', technologyID, 'patents', technologyID), stripPayableDetail

    @store.fetchAll
    def getPayableDetailsForPatent(self, patentID):
        'Get payableDetails related to patent'
        sql = buildPayableDetailSQL(where='paybldtl.linktable=? AND paybldtl.linkfk=?') + ' ORDER BY payable.duedate DESC'
        return sql, ('patents', patentID), stripPayableDetail

    # Remark

    def addRemark(self, linkTable, linkFK, remark, remarkDate):
        'Add a remark'
        return self.addRemarkFromDictionary({
            'linktable': linkTable,
            'linkfk': linkFK,
            'remark': remark,
            'remarkdt': remarkDate,
        })

    def addRemarkFromDictionary(self, valueByName):
        'Add a remark'
        # Initialize
        now = makeNow()
        # Add
        self.add('remarks', [
            ('linktable', '', True),
            ('linkfk', 0, True),
            ('orderno', self.getNextRemarkOrder(valueByName['linktable'], valueByName['linkfk']), False),
            ('alinkfk', 50484, False),
            ('alinktable', 'contacts', False),
            ('private', 'N', False),
            ('remark', '', True),
            ('remarkrtf', makeRTF(valueByName['remark']) if 'remark' in valueByName else '', False),
            ('remarkdt', now, False),
        ], valueByName)
        # Prepare
        remarkType = {'agrmnts': 'Agreement', 'company': 'Company', 'contacts': 'Contact', 'mktprj': 'Marketing Project', 'mkttgt': 'Marketing Target', 'patents': 'Patent', 'payable': 'Payable', 'recvable': 'Receivable', 'technol': 'Technology'}[valueByName['linktable']] 
        # Add audit
        self.addAudit('remarks', valueByName['primarykey'], 'Type: %s, By: %s' % (remarkType, 'Admin'), '[OrderNo] [ActivityFK] [LinkFK] [LinkTable] [ALinkFK] [ALinkTable] [RemarkDt] [Type] [Subject] [EntryID] [RcvdDate] [SentDate] [Remark] [RemarkRTF]')
        # Commit
        self.connection2.commit()
        # Return
        return valueByName

    @store.commit
    def deleteRemark(self, remarkID):
        'Delete remark'
        return 'DELETE FROM remarks WHERE primarykey=?', (remarkID,)

    def findRemarkIDs(self, term):
        'Return matching remarks'
        # Return
        return self.findIDs('remarks', [], [], ['remark LIKE ?'], term, ['remarks.remarkdt DESC'])

    def getRemarkBundle(self, remarkID, simple=False):
        'Get information related to the remark'
        # Set value
        value = [remarkID]
        # Get remark
        self.cursor1.execute(buildRemarkSQL(where='remarks.primarykey=?'), value)
        remark = stripRemark(self.cursor1.fetchone())
        if simple:
            return remark
        # Get remarkLink
        remarkLink = self.getLink(remark['linkTable'], remark['linkID'])
        # Return
        return remark, remarkLink

    def getRemarks(self, linkTable, linkFK):
        'Get corresponding remarks'
        # Get remarks
        self.cursor1.execute(buildRemarkSQL(where='linktable=? AND linkfk=?') + ' ORDER BY remarkdt DESC', (linkTable, linkFK))
        # Return
        return map(stripRemark, self.cursor1.fetchall())

    # Technology

    def addTechnology(self, technologyTitle):
        'Add a technology'
        return self.addTechnologyFromDictionary({
            'name': technologyTitle,
        })

    def addTechnologyFromDictionary(self, valueByName):
        'Add a technology'
        # Initialize
        now = makeNow()
        # Add
        self.add('technol', [
            ('techid', self.getNextCaseNumber(), False),
            ('name', '', False),
            ('managerfk', 1044, False),
            ('discldt', now, False),
        ], valueByName)
        # Add audit
        self.addAudit('technol', valueByName['primarykey'], 'Tech ID:%s' % valueByName['techid'], '[Name] [ManagerFK] [TechId] [AssignCliD] [DisclDt] [AssignInvD] [AssignMgrD] [PubDisDt] [StatUpdate] [ProjectsFK] [TechStatFK] [RDBalance] [UDF_keu] [UDF_LF] [UDF_PatApp] [UDF_Issued] [UDF_MktRsc] [UDF_Licens] [UDF_Refnum] [UDF_Keyw]')
        # Commit
        self.connection2.commit()
        # Return
        return valueByName

    def findTechnologyIDs(self, term, contactIDString=None):
        'Return matching technologyIDs'
        # Initialize
        joins, wheres = set(), set()
        # Use contacts
        if contactIDString == None:
            contactIDString = makeIDString(self.findContactIDs(term))
        if contactIDString:
            joins.add('LEFT JOIN techinv ON technol.primarykey=techinv.technolfk')
            wheres.add('techinv.contactsfk IN (%s)' % contactIDString)
        # Build patterns
        patterns = [
            'technol.name LIKE ?', 
            'technol.techid LIKE ?',
        ]
        # Return
        return self.findIDs('technol', joins, wheres, patterns, term, ['technol.techid'])

    def getNextCaseNumber(self):
        """
        Get next available case number for a technology.
        Note that there will be a problem with this numbering 
        scheme at the start of the next millennium.
        """
        # Get now
        now = datetime.date.today()
        # Get fiscal year
        year = now.year + 1 if now.month > 6 else now.year
        # Get last two digits
        yearDigits = ('%s' % year)[-2:]
        # Get case numbers beginning with the last two digits of this year
        value = [yearDigits + 'A%']
        self.cursor1.execute('SELECT TechId FROM TECHNOL WHERE TechId LIKE ?', value)
        caseNumbers = [x[0].strip() for x in self.cursor1.fetchall()]
        # Extract last four digits
        fourDigits = [int(x[-4:]) for x in caseNumbers]
        # Get largest last four digits
        largestFourDigits = max(fourDigits)
        # Return
        return '%sA%04d' % (yearDigits, largestFourDigits + 1)

    def getTechnologyBundle(self, technologyID, simple=False):
        'Get information related to the technology'
        # Set value
        value = [technologyID]
        # Get technology
        self.cursor1.execute(buildTechnologySQL(where='technol.primarykey=?'), value)
        technology = stripTechnology(self.cursor1.fetchone())
        if simple:
            return technology
        # Get technologyContactBundles
        self.cursor1.execute('SELECT contacts.primarykey FROM contacts INNER JOIN techinv ON contacts.primarykey=techinv.contactsfk WHERE techinv.technolfk=? ORDER BY pi_order', value)
        technologyContactBundles = [self.getContactBundle(x[0]) for x in self.cursor1.fetchall()]
        technology['leadInventor'] = message_format.formatContactName(technologyContactBundles[0][0]) if technologyContactBundles else ''
        # Get technologyRemarks
        technologyRemarks = self.getRemarks('technol', technologyID)
        # Get technologyPayableDetails
        technologyPayableDetails = self.getPayableDetailsForTechnology(technologyID)
        # Get technologyPatentBundles
        self.cursor1.execute('SELECT primarykey FROM patents WHERE technolfk=? ORDER BY patents.filedate', value)
        technologyPatentBundles = [self.getPatentBundle(x[0]) for x in self.cursor1.fetchall()]
        # Get technologyAgreementBundles
        self.cursor1.execute('SELECT agrmnts.primarykey FROM agrmnts INNER JOIN agrtech ON agrmnts.primarykey=agrtech.agrmntsfk WHERE agrtech.technolfk=? ORDER BY agrmnts.agrmntid', value)
        technologyAgreementBundles = [self.getAgreementBundle(x[0]) for x in self.cursor1.fetchall()]
        # Return
        return technology, technologyContactBundles, technologyRemarks, technologyPayableDetails, technologyPatentBundles, technologyAgreementBundles


# Make

def makeIDString(ids): 
    'Convert ids into a comma-joined string'
    return ','.join(str(x) for x in ids)

def makeNow():
    'Return date as a string'
    return datetime.date.today().strftime('%m/%d/%Y')

def makeRTF(text):
    'Convert plain text into RTF format'
    return '{\\rtf1\\ansi\\ansicpg1252\\deff0\\deflang1033{\\fonttbl{\\f0\\fnil\\fcharset0 MS Shell Dlg 2;}}\n\\viewkind4\\uc1\\pard\\f0\\fs17\n%s\n\\par\n}\n' % text.replace('\n', '\n\\par ')


# Build

def buildAgreementSQL(join='', where=''):
    whereString = 'WHERE ' + where if where else ''
    return 'SELECT agrmnts.primarykey, agrmntid, name FROM agrmnts %s %s ORDER BY agrmnts.agrmntid' % (join, whereString)

def buildCompanySQL(join='', where=''):
    whereString = 'WHERE ' + where if where else ''
    return 'SELECT company.primarykey, name, phone, website FROM company %s %s ORDER BY name' % (join, whereString)

def buildContactSQL(join='', where=''):
    whereString = 'WHERE ' + where if where else ''
    return 'SELECT contacts.primarykey, firstname, middleini, lastname, email, company.name, contacts.address1, contacts.address2, contacts.address3, contacts.city, contacts.stateprv, contacts.postalcode, country.name, contacts.comments FROM contacts LEFT JOIN company ON contacts.companyfk=company.primarykey LEFT JOIN country ON contacts.countryfk=country.primarykey %s %s ORDER BY lastname, firstname, middleini' % (join, whereString)

def buildDocumentSQL(join='', where=''):
    whereString = 'WHERE ' + where if where else ''
    return 'SELECT document.primarykey, name, linkfk, linktable, fileexten FROM document %s %s ORDER BY name' % (join, whereString)

def buildMarketingProjectSQL(join='', where=''):
    whereString = 'WHERE ' + where if where else ''
    return 'SELECT mktprj.primarykey, mktprjid, mktprj.name, mkprstat.name FROM mktprj LEFT JOIN mkprstat ON mktprj.mkprstatfk=mkprstat.primarykey %s %s ORDER BY mktprjid' % (join, whereString)

def buildMarketingTargetSQL(join='', where=''):
    whereString = 'WHERE ' + where if where else ''
    return 'SELECT mkttgt.primarykey, linktable, linkfk, company.name, contacts.firstname, contacts.middleini, contacts.lastname, mktgstat.name FROM mkttgt LEFT JOIN company ON mkttgt.companyfk=company.primarykey LEFT JOIN contacts ON mkttgt.contactsfk=contacts.primarykey LEFT JOIN mktgstat ON mkttgt.mktgstatfk=mktgstat.primarykey %s %s ORDER BY company.name' % (join, whereString)

def buildPatentSQL(join='', where=''):
    whereString = 'WHERE ' + where if where else ''
    return 'SELECT patents.primarykey, technol.techid, patents.name, papptype.name, company.name, patents.legalrefno, patents.filedate, patents.serialno, patents.patentno FROM patents LEFT JOIN company ON company.primarykey=patents.lawfirmfk LEFT JOIN papptype ON papptype.primarykey=patents.papptypefk LEFT JOIN technol ON technol.primarykey=patents.technolfk %s %s ORDER BY technol.techid, patents.filedate' % (join, whereString)

def buildPayableSQL(join='', where=''):
    whereString = 'WHERE ' + where if where else ''
    return 'SELECT payable.primarykey, payable.invoiceno, payable.invoicedt, payable.duedate FROM payable %s %s' % (join, whereString)

def buildPayableDetailSQL(join='', where=''):
    whereString = 'WHERE ' + where if where else ''
    return 'SELECT paybldtl.primarykey, paybldtl.payablefk, payable.invoiceno, paybldtl.amount, payable.invoicedt, payable.duedate FROM paybldtl INNER JOIN payable ON paybldtl.payablefk=payable.primarykey %s %s' % (join, whereString)

def buildPaymentSQL(join='', where=''):
    whereString = 'WHERE ' + where if where else ''
    return 'SELECT payment.primarykey, linkfk, linktable, pmtamount, pmtdate, checknum FROM payment %s %s' % (join, whereString)

def buildReceivableSQL(join='', where=''):
    whereString = 'WHERE ' + where if where else ''
    return 'SELECT recvable.primarykey, agrparty.linkfk, agrparty.linktable, invoiceno, invoicedt, duedate FROM recvable INNER JOIN agrparty ON recvable.agrpartyfk=agrparty.primarykey %s %s' % (join, whereString)

def buildRemarkSQL(join='', where=''):
    whereString = 'WHERE ' + where if where else ''
    return 'SELECT remarks.primarykey, remarks.remarkdt, remarks.remark, remarks.linktable, remarks.linkfk FROM remarks %s %s' % (join, whereString)

def buildTechnologySQL(join='', where=''):
    whereString = 'WHERE ' + where if where else ''
    return 'SELECT technol.primarykey, technol.techid, technol.name, projects.name, techstat.name, technol.discldt, technol.comments FROM technol LEFT JOIN projects ON technol.projectsfk=projects.primarykey LEFT JOIN techstat ON technol.techstatfk=techstat.primarykey %s %s ORDER BY techid' % (join, whereString)


# Strip

def stripAgreement(result):
    'Strip whitespace from agreement fields and return a dictionary'
    agreementID, agreementCaseNumber, agreementTitle = map(store.escape, result)
    return dict(agreementID=agreementID, agreementCaseNumber=agreementCaseNumber, agreementTitle=agreementTitle)

def stripCompany(result):
    'Strip whitespace from company fields and return a dictionary'
    companyID, companyName, companyPhone, companyWebsite = map(store.escape, result)
    return dict(companyID=companyID, companyName=companyName, companyPhone=companyPhone, companyWebsite=companyWebsite)

def stripContact(result):
    'Strip whitespace from contact fields and return a dictionary'
    contactID, firstName, middleName, lastName, email, company, address1, address2, address3, city, state, postalCode, country, comments = map(store.escape, result)
    return dict(contactID=contactID, email=email, firstName=firstName, middleName=middleName, lastName=lastName, address1=address1, address2=address2, address3=address3, city=city, state=state, postalCode=postalCode, country=country, company=company, comments=comments)

def stripDocument(result):
    'Strip whitespace from document fields and return a dictionary'
    documentID, documentName, linkID, linkTable, fileExtension = map(store.escape, result)
    return dict(documentID=documentID, documentName=documentName, linkID=linkID, linkTable=linkTable.lower(), fileExtension=fileExtension.lower())

def stripMarketingProject(result):
    'Strip whitespace from marketing project fields and return a dictionary'
    marketingProjectID, marketingProjectCaseNumber, marketingProjectName, marketingProjectStatus = map(store.escape, result)
    return dict(marketingProjectID=marketingProjectID, marketingProjectCaseNumber=marketingProjectCaseNumber, marketingProjectName=marketingProjectName, marketingProjectStatus=marketingProjectStatus)

def stripMarketingTarget(result):
    'Strip whitespace from marketing target fields and return a dictionary'
    marketingTargetID, linkTable, linkID, companyName, firstName, middleName, lastName, marketingTargetStatus = map(store.escape, result)
    return dict(marketingTargetID=marketingTargetID, linkTable=linkTable.lower(), linkID=linkID, companyName=companyName, contactName=message_format.formatContact(dict(firstName=firstName, middleName=middleName, lastName=lastName)), marketingTargetStatus=marketingTargetStatus)

def stripPatent(result):
    'Strip whitespace from patent fields and return a dictionary'
    patentID, patentCaseNumber, patentTitle, patentApplicationType, lawFirm, legalReferenceNumber, patentFilingDate, patentSerialNumber, patentNumber = map(store.escape, result)
    return dict(patentID=patentID, patentCaseNumber=patentCaseNumber, patentTitle=patentTitle, patentApplicationType=patentApplicationType, lawFirm=lawFirm, legalReferenceNumber=legalReferenceNumber, patentFilingDate=store.formatDate(patentFilingDate), patentSerialNumber=patentSerialNumber, patentNumber=patentNumber)

def stripPayable(result):
    payableID, invoiceNumber, invoiceDate, invoiceDateDue = map(store.escape, result)
    return dict(payableID=payableID, invoiceNumber=invoiceNumber, invoiceDate=invoiceDate, invoiceDateDue=invoiceDateDue)

def stripPayableDetail(result):
    'Strip whitespace from payableDetail fields and return a dictionary'
    payableDetailID, payableID, invoiceNumber, amount, invoiceDate, invoiceDateDue = map(store.escape, result)
    return dict(payableDetailID=payableDetailID, payableID=payableID, invoiceNumber=invoiceNumber, amount='$ %.02f' % amount, invoiceDate=store.formatDate(invoiceDate), invoiceDateDue=store.formatDate(invoiceDateDue))

def stripPayment(result):
    'Strip whitespace from payment fields and return a dictionary'
    paymentID, linkID, linkTable, paymentAmount, paymentDate, checkNumber = map(store.escape, result)
    return dict(paymentID=paymentID, linkID=linkID, linkTable=linkTable.lower(), paymentAmount=paymentAmount, paymentDate=paymentDate, checkNumber=checkNumber)

def stripPhone(result):
    'Strip whitespace from phone fields and return a dictionary'
    phoneID, phoneNumber, phoneType = map(store.escape, result)
    return dict(phoneID=phoneID, phoneNumber=phoneNumber, phoneType=phoneType)

def stripReceivable(result):
    'Strip whitespace from receivable fields and return a dictionary'
    receivableID, linkID, linkTable, invoiceNumber, invoiceDate, invoiceDateDue = map(store.escape, result)
    return dict(receivableID=receivableID, linkID=linkID, linkTable=linkTable.lower(), invoiceNumber=invoiceNumber, invoiceDate=invoiceDate, invoiceDateDue=invoiceDateDue)

def stripRemark(result):
    'Strip whitespace from remark fields and return a dictionary'
    remarkID, remarkDate, remark, linkTable, linkID = map(store.escape, result)
    return dict(remarkID=remarkID, remarkDate=store.formatDate(remarkDate), remark=remark.replace('\r', ''), linkTable=linkTable.lower(), linkID=linkID)

def stripTechnology(result):
    'Strip whitespace from technology fields and return a dictionary'
    technologyID, technologyCaseNumber, technologyTitle, projectName, technologyStatus, disclosureDate, technologyComments = map(store.escape, result)
    return dict(technologyID=technologyID, technologyCaseNumber=technologyCaseNumber, technologyTitle=technologyTitle, disclosureDate=store.formatDate(disclosureDate), projectName=projectName, technologyStatus=technologyStatus, technologyComments=technologyComments)


# Exception

class ExtractionError(Exception):
    'Exception that indicates problems during value extraction'
    pass
