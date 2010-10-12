'Routines and constants for formatting replies'
# Import system modules
import re
# Import custom modules
import store
import config


# Set templates
templateOriginal = 'From: %s\nTo: %s\nDate: %s\nSubject: %s\n\n%s'
templateTable = '<table>\n<tr>\n%s\n</tr>\n</table>'
# Load templates for documents
templateErrorFatal = store.loadTemplate('ErrorFatal.html')
templateErrorSyntax = store.loadTemplate('ErrorSyntax.html')
templateHelp = store.loadTemplate('Help.html') % {
    'address': config.mailAddress, 
}
templateQuery = store.loadTemplate('Query.html')
templateQueryResult = store.loadTemplate('QueryResult.html')
templateRemarkResult = store.loadTemplate('RemarkResult.html')
templateRemarkFailed = store.loadTemplate('RemarkFailed.html')
templateReservation = store.loadTemplate('Reservation.html')
templateRetrieval = store.loadTemplate('Retrieval.html')
# Load templates for results
templateAgreementInformation = """\
<table class=information>
    <tr>
        <td class=informationField>Case Number</td>
        <td>%(agreementCaseNumber)s</td>
    </tr>
    <tr>
        <td class=informationField>Title</td>
        <td>%(agreementTitle)s</td>
    </tr>
</table>"""
templateAgreementLabel = """\
        <td class=caseNumber>%(agreementCaseNumber)s</td>
        <td class=title>%(agreementTitle)s</td>"""
templateCompanyInformation = """\
<table class=information>
    <tr>
        <td class=informationField>Company Name</td>
        <td>%(companyName)s</td>
    </tr>
    %(companyInformation)s
</table>"""
templateCompanyLabel = """\
        <td class=companyName>%(companyName)s</td>"""
templateContactInformation = """\
<table class=information>
    <tr>
        <td class=informationField>Contact Name</td>
        <td>%(contactName)s</td>
    </tr>
    %(contactInformation)s
</table>"""
templateContactLabel = """\
        <td class=contactName>%(lastName)s, %(firstName)s %(middleName)s</td>"""
templateDocumentInformation = """\
<table class=information>
    <tr>
        <td class=informationField>Document ID</td>
        <td>%(documentID)s</td>
    </tr>
    <tr>
        <td class=informationField>Document Name</td>
        <td>%(documentName)s</td>
    </tr>
    <tr>
        <td class=informationField>File Extension</td>
        <td>%(fileExtension)s</td>
    </tr>
    <tr>
        <td class=informationField>Link</td>
        <td>%(linkLabel)s</td>
    </tr>
</table>"""
templateDocumentLabel = """\
        <td class=caseNumber>%(documentID)s</td>
        <td class=title>%(documentName)s</td>
        <td>%(fileExtension)s</td>"""
templateMarketingProjectInformation = """\
<table class=information>
    <tr>
        <td class=informationField>Case Number</td>
        <td>%(marketingProjectCaseNumber)s</td>
    </tr>
    <tr>
        <td class=informationField>Name</td>
        <td>%(marketingProjectName)s</td>
    </tr>
    <tr>
        <td class=informationField>Status</td>
        <td>%(marketingProjectStatus)s</td>
    </tr>
</table>"""
templateMarketingProjectLabel = """\
        <td class=caseNumber>%(marketingProjectCaseNumber)s</td>
        <td class=title>%(marketingProjectName)s</td>
        <td class=status>%(marketingProjectStatus)s</td>"""
templateMarketingTargetInformation = """\
<table class=information>
    <tr>
        <td class=informationField>Target Company</td>
        <td>%(companyName)s</td>
    </tr>
    <tr>
        <td class=informationField>Target Contact</td>
        <td>%(contactName)s</td>
    </tr>
    <tr>
        <td class=informationField>Status</td>
        <td>%(marketingTargetStatus)s</td>
    </tr>
    <tr>
        <td class=informationField>Link</td>
        <td>%(linkLabel)s</td>
    </tr>
</table>"""
templateMarketingTargetLabel = """\
        <td class=companyName>%(companyName)s</td>
        <td class=contactName>%(contactName)s</td>
        <td class=status>%(marketingTargetStatus)s</td>"""
templatePatentInformation = """\
<table class=information>
    <tr>
        <td class=informationField>Case Number</td>
        <td>%(patentCaseNumber)s</td>
    </tr>
    <tr>
        <td class=informationField>Law Firm</td>
        <td>%(lawFirm)s</td>
    </tr>
    <tr>
        <td class=informationField>Legal Reference Number</td>
        <td>%(legalReferenceNumber)s</td>
    </tr>
    <tr>
        <td class=informationField>Title</td>
        <td>%(patentTitle)s</td>
    </tr>
    <tr>
        <td class=informationField>Application Type</td>
        <td>%(patentApplicationType)s</td>
    </tr>
    <tr>
        <td class=informationField>Application Filing Date</td>
        <td>%(patentFilingDate)s</td>
    </tr>
    <tr>
        <td class=informationField>Application Serial Number</td>
        <td>%(patentSerialNumber)s</td>
    </tr>
    <tr>
        <td class=informationField>Patent Number</td>
        <td>%(patentNumber)s</td>
    </tr>
</table>"""
templatePatentLabel = """\
        <td class=caseNumber>%(patentCaseNumber)s</td>
        <td class=title>%(patentTitle)s</td>
        <td class=date>%(patentFilingDate)s</td>
        <td class=patentApplicationType>%(patentApplicationType)s</td>
        <td class=patentSerialNumber>%(patentSerialNumber)s</td>
        <td class=leadInventor>%(leadInventor)s</td>"""
templatePayableLabel = """\
        <td class=caseNumber>%(invoiceNumber)s</td>
        <td class=date>%(invoiceDate)s</td>
        <td class=date>%(invoiceDateDue)s</td>"""
templatePayableDetailLabel = """\
        <td class=date>%(invoiceDate)s</td>
        <td class=amount>%(amount)s</td>"""
templatePayableDetailSubstance = """\
<table class=information>
    <tr>
        <td class=informationField>Invoice Number</td>
        <td>%(invoiceNumber)s</td>
    </tr>
    <tr>
        <td class=informationField>Invoice Date</td>
        <td>%(invoiceDate)s</td>
    </tr>
    <tr>
        <td class=informationField>Due Date</td>
        <td>%(invoiceDateDue)s</td>
    </tr>
    <tr>
        <td class=informationField>Amount</td>
        <td>%(amount)s</td>
    </tr>
</table>"""
templatePaymentLabel = """\
        <td class=caseNumber>%(checkNumber)s</td>
        <td class=amount>%(paymentAmount)s</td>
        <td class=date>%(paymentDate)s</td>"""
templateReceivableLabel = """\
        <td class=caseNumber>%(invoiceNumber)s</td>
        <td class=date>%(invoiceDate)s</td>
        <td class=date>%(invoiceDateDue)s</td>"""
templateRemarkInformation = """\
<table class=information>
    <tr>
        <td class=informationField>Remark Date</td>
        <td>%(remarkDate)s</td>
    </tr>
    <tr>
        <td class=informationField>Link</td>
        <td>%(linkLabel)s</td>
    </tr>
</table>"""
templateRemarkLabel = """\
        <td class=date>%(remarkDate)s</td>
        <td>%(remarkSummary)s</td>"""
templateTechnologyInformation = """\
<table class=information>
    <tr>
        <td class=informationField>Case Number</td>
        <td>%(technologyCaseNumber)s</td>
    </tr>
    <tr>
        <td class=informationField>Title</td>
        <td>%(technologyTitle)s</td>
    </tr>
    <tr>
        <td class=informationField>Status</td>
        <td>%(technologyStatus)s</td>
    </tr>
    <tr>
        <td class=informationField>Project</td>
        <td>%(projectName)s</td>
    </tr>
    <tr>
        <td class=informationField>Lead Inventor</td>
        <td>%(leadInventor)s</td>
    </tr>
    <tr>
        <td class=informationField>Disclosure Date</td>
        <td>%(disclosureDate)s</td>
    </tr>
</table>"""
templateTechnologyLabel = """\
        <td class=caseNumber>%(technologyCaseNumber)s</td>
        <td class=title>%(technologyTitle)s</td>
        <td class=status>%(technologyStatus)s</td>
        <td class=project>%(projectName)s</td>
        <td class=leadInventor>%(leadInventor)s</td>"""


# Core

def formatErrorFatal(message, error):
    return templateErrorFatal % {
        'originalEmail': formatMailForward(message),
        'error': error,
    }

def formatErrorSyntax(message):
    return formatHelp(templateErrorSyntax % {
        'originalEmail': formatMailForward(message),
    })

def formatHelp(message=''):
    return templateHelp % {
        'message': message,
    }

def formatQueryResults(label, resultPacks, terms):
    # Initialize
    outlines = []
    contents = []
    formatBundle, formatRow = formatPackByLabel[label]
    # For each resultPack,
    for resultIndex, (termsIncluded, bundles) in enumerate(resultPacks):
        # Start
        outlines.append('<div class=header>%s</div>' % ' '.join(x if x in termsIncluded else '<s>%s</s>' % x for x in terms))
        outlines.append('<table>')
        # For each bundle,
        for bundleIndex, bundle in enumerate(bundles):
            # Prepare
            sectionLink = 'section%s-%s' % (resultIndex, bundleIndex)
            # Format
            sectionRow = formatRow(bundle[0])
            sectionContent = formatBundle(bundle)
            # Append
            outlines.append('<tr id=%s class=tabOFF onclick=go(this) onmouseover=on(this) onmouseout=off(this)>%s</tr>' % (sectionLink, sectionRow))
            contents.append('<a name=%s_></a>' % sectionLink)
            contents.append('<div id=%s_ class=tab_>%s</div>' % (sectionLink, sectionContent))
        # End
        outlines.append('</table>')
    # Return
    return templateQueryResult % {
        'title': label,
        'outline': '\n'.join(outlines),
        'content': '\n'.join(contents),
    }

def formatRemarkResults(message, summaries, attachmentNames):
    # Count
    attachmentCount = len(attachmentNames)
    if attachmentCount == 0:
        attachmentMessage = ''
    elif attachmentCount == 1:
        attachmentMessage = 'and its attachment as a document'
    else:
        attachmentMessage = 'and its attachments as documents'
    # Set
    return templateRemarkResult % {
        'attachmentMessage': attachmentMessage,
        'pluralEnding': 's' if len(summaries) > 1 else '',
        'summaries': '<br>\n'.join(summaries),
        'originalEmail': formatMailForward(message),
        'documents': '<b>Documents</b><br>\n' + '<br>'.join(attachmentNames) if attachmentCount else '',
    }

def formatRemarkFailed(message, error):
    return templateRemarkFailed % {
        'originalEmail': formatMailForward(message),
        'error': error,
    }


# Mail

def formatMailOriginal(message):
    return templateOriginal % (message['from'], message['to'], store.formatDateTime(message['when']), message['subject'], message['body'])

def formatMailForward(message):
    return indentText(formatMailOriginal(message), '> ')

def indentText(text, prefix):
    return '\n'.join(prefix + line for line in text.splitlines())


# Bundle

def formatAgreementBundle(agreementBundle, linkPrefix=''):
    'Format agreementBundle'
    # Initialize
    agreement, agreementResearcherContactBundles, agreementPartyContactBundles, agreementPartyCompanyBundles, agreementRemarks, agreementTechnologies, agreementPatents = agreementBundle
    link = linkPrefix + '-agreement%(agreementID)s' % agreement
    contents = []
    # Format information
    contents.append(templateAgreementInformation % agreement)
    # Format
    contents.append(formatContactBundles(agreementResearcherContactBundles, link, 'Researchers'))
    contents.append(formatContactBundles(agreementPartyContactBundles, link + '-parties', 'People'))
    contents.append(formatCompanyBundles(agreementPartyCompanyBundles, link + '-parties', 'Companies'))
    contents.append(formatRemarks(agreementRemarks, link))
    if agreementTechnologies:
        contents.append(formatFold(link + '-technologies', formatLabel(agreementTechnologies, 'Technologies'), templateTable % '</tr>\n<tr>'.join(map(formatTechnology, agreementTechnologies))))
    if agreementPatents:
        contents.append(formatFold(link + '-patents', formatLabel(agreementPatents, 'Patents'), templateTable % '</tr>\n<tr>'.join(map(formatPatent, agreementPatents))))
    # Return
    return formatFold(link, templateTable % formatAgreement(agreement), '\n'.join(contents)) if linkPrefix else '\n'.join(contents)

def formatCompanyBundle(companyBundle, linkPrefix=''):
    'Format companyBundle'
    # Initialize
    company, companyContactBundles, companyRemarks = companyBundle
    link = linkPrefix + '-company%(companyID)s' % company
    contents = []
    # Format information
    informations = []
    if company['companyPhone']:
        phoneTemplate = '<td class=informationField>Phone</td><td>%(companyPhone)s</td>'
        informations.append(phoneTemplate % company)
    if company['companyWebsite']:
        websiteTemplate = '<td class=informationField>Website</td><td><a href="%(companyWebsite)s">%(companyWebsite)s</a></td>'
        informations.append(websiteTemplate % company)
    # Format
    contents.append(templateCompanyInformation % {
        'companyName': company['companyName'],
        'companyInformation': '<tr>%s</tr>' % '</tr><tr>'.join(informations),
    })
    contents.append(formatContactBundles(companyContactBundles, link, 'Contacts'))
    contents.append(formatRemarks(companyRemarks, link))
    # Return
    return formatFold(link, templateTable % formatCompany(company), '\n'.join(contents)) if linkPrefix else '\n'.join(contents)

def formatContactBundle(contactBundle, linkPrefix=''):
    'Format contactBundle'
    # Initialize
    contact, contactPhones, contactRemarks = contactBundle
    link = linkPrefix + '-contact%(contactID)s' % contact
    contents = []
    # Format information
    informations = []
    if contact['email']:
        emailTemplate = '<a href="mailto:%(email)s">%(email)s</a>'
        informations.append('<td class=informationField>Email</td><td>%s</td>' % emailTemplate % contact)
    for phone in contactPhones:
        if '@' in phone['phoneNumber']:
            phoneTemplate = '<a href="mailto:%(phoneNumber)s">%(phoneNumber)s</a>'
        else: 
            phoneTemplate = '%(phoneNumber)s (%(phoneType)s)'
        informations.append('<td class=informationField>Phone</td><td>%s</td>' % (phoneTemplate % phone))
    address = patternWhitespace.sub(' ', '%(address1)s %(address2)s %(address3)s %(city)s %(state)s %(postalCode)s %(country)s' % contact)
    if address: 
        informations.append('<td class=informationField>Address</td><td>%s</td>' % address)
    if contact['company']: 
        informations.append('<td class=informationField>Phone</td><td>%s</td>' % contact['company'])
    # Format
    contents.append(templateContactInformation % {
        'contactName': formatContactName(contact),
        'contactInformation': '<tr>%s</tr>' % '</tr><tr>'.join(informations),
    })
    contents.append(formatFold(link + '-comments', 'Comments', contact['comments'].replace('\n', '<br>\n')))
    contents.append(formatRemarks(contactRemarks, link))
    # Return
    return formatFold(link, templateTable % formatContact(contact), '\n'.join(contents)) if linkPrefix else '\n'.join(contents)

def formatDocumentBundle(documentBundle, linkPrefix=''):
    'Format documentBundle'
    # Initialize
    document, documentLink = documentBundle
    link = linkPrefix + '-document%(documentID)s' % document
    contents = []
    document['linkLabel'] = label = labelByTable[document['linkTable']]
    linkFormat = formatPackByLabel[label][1]
    # Format information
    contents.append(templateDocumentInformation % document)
    # Format
    contents.append(templateTable % linkFormat(documentLink))
    # Return
    return formatFold(link, templateTable % formatDocument(document), '\n'.join(contents)) if linkPrefix else '\n'.join(contents)

def formatMarketingProjectBundle(marketingProjectBundle, linkPrefix=''):
    'Format marketingProjectBundle'
    # Initialize
    marketingProject, marketingProjectMarketingTargetBundles, marketingProjectContactBundles, marketingProjectRemarks, marketingProjectTechnologies, marketingProjectPatents = marketingProjectBundle
    link = linkPrefix + '-marketingproject%(marketingProjectID)s' % marketingProject
    contents = []
    # Format information
    contents.append(templateMarketingProjectInformation % marketingProject)
    # Format
    contents.append(formatMarketingTargetBundles(marketingProjectMarketingTargetBundles, link))
    contents.append(formatContactBundles(marketingProjectContactBundles, link, 'Inventors'))
    contents.append(formatRemarks(marketingProjectRemarks, link))
    if marketingProjectTechnologies:
        contents.append(formatFold(link + '-technologies', formatLabel(marketingProjectTechnologies, 'Technologies'), templateTable % '</tr>\n<tr>'.join(map(formatTechnology, marketingProjectTechnologies))))
    if marketingProjectPatents:
        contents.append(formatFold(link + '-patents', formatLabel(marketingProjectPatents, 'Patents'), templateTable % '</tr>\n<tr>'.join(map(formatPatent, marketingProjectPatents))))
    # Return
    return formatFold(link, templateTable % formatMarketingProject(marketingProject), '\n'.join(contents)) if linkPrefix else '\n'.join(contents)

def formatMarketingTargetBundle(marketingTargetBundle, linkPrefix=''):
    'Format marketingTargetBundle'
    # Initialize
    marketingTarget, marketingTargetLink, marketingTargetRemarks, marketingTargetAgreements = marketingTargetBundle
    link = linkPrefix + '-marketingTarget%(marketingTargetID)s' % marketingTarget
    contents = []
    marketingTarget['linkLabel'] = label = labelByTable[marketingTarget['linkTable']]
    linkFormat = formatPackByLabel[label][1]
    # Format information
    contents.append(templateMarketingTargetInformation % marketingTarget)
    # Format
    contents.append(templateTable % linkFormat(marketingTargetLink))
    contents.append(formatRemarks(marketingTargetRemarks, link))
    if marketingTargetAgreements:
        contents.append(formatFold(link + '-agreements', formatLabel(marketingTargetAgreements, 'Agreements'), templateTable % '</tr>\n<tr>'.join(map(formatAgreement, marketingTargetAgreements))))
    # Return
    return formatFold(link, templateTable % formatMarketingTarget(marketingTarget), '\n'.join(contents)) if linkPrefix else '\n'.join(contents)

def formatPatentBundle(patentBundle, linkPrefix=''):
    'Format patentBundle'
    # Initialize
    patent, patentContactBundles, patentRemarks, patentPayableDetails = patentBundle
    link = linkPrefix + '-patent%(patentID)s' % patent
    contents = []
    # Format information
    contents.append(templatePatentInformation % patent)
    # Format
    contents.append(formatContactBundles(patentContactBundles, link, 'Inventors'))
    contents.append(formatRemarks(patentRemarks, link))
    contents.append(formatPayableDetails(patentPayableDetails, link))
    # Return
    return formatFold(link, templateTable % formatPatent(patent), '\n'.join(contents)) if linkPrefix else '\n'.join(contents)

# def formatPayableBundle(payableBundle, linkPrefix=''):
    # 'Format payableBundle'
    # return ''

# def formatReceivableBundle(receivableBundle, linkPrefix=''):
    # 'Format receivableBundle'
    # return ''

def formatRemarkBundle(remarkBundle, linkPrefix=''):
    'Format remarkBundle'
    # Initialize
    remark, remarkLink = remarkBundle
    link = linkPrefix + '-remark%(remarkID)s' % remark
    contents = []
    remark['linkLabel'] = label = labelByTable[remark['linkTable']]
    linkFormat = formatPackByLabel[label][1]
    # Format information
    contents.append(templateRemarkInformation % remark)
    # Format
    contents.append(templateTable % linkFormat(remarkLink))
    contents.append('<div class=remark>%s</div>' % remark['remark'].replace('\n', '<br>'))
    # Return
    return formatFold(link, templateTable % formatRemark(remark), '\n'.join(contents)) if linkPrefix else '\n'.join(contents)

def formatTechnologyBundle(technologyBundle, linkPrefix=''):
    'Format technologyBundle'
    # Initialize
    technology, technologyContactBundles, technologyRemarks, technologyPayableDetails, technologyPatentBundles, technologyAgreementBundles = technologyBundle
    link = linkPrefix + '-technology%(technologyID)s' % technology
    contents = []
    # Format information
    contents.append(templateTechnologyInformation % technology)
    # Format
    contents.append(formatFold(link + '-ncd', 'Non-Confidential Disclosure', technology['technologyComments']))
    contents.append(formatContactBundles(technologyContactBundles, link, 'Inventors'))
    contents.append(formatRemarks(technologyRemarks, link))
    contents.append(formatPayableDetails(technologyPayableDetails, link))
    contents.append(formatPatentBundles(technologyPatentBundles, link))
    contents.append(formatAgreementBundles(technologyAgreementBundles, link))
    # Return
    return formatFold(link, templateTable % formatTechnology(technology), '\n'.join(contents)) if linkPrefix else '\n'.join(contents)


# Bundles

def formatBundles(bundles, linkPrefix, label, formatBundle):
    return formatFold(linkPrefix + label.lower().replace(' ', ''), formatLabel(bundles, label), '\n'.join(formatBundle(x, linkPrefix) for x in bundles))

def formatAgreementBundles(bundles, linkPrefix):
    return formatBundles(bundles, linkPrefix, 'Agreements', formatAgreementBundle)

def formatCompanyBundles(bundles, linkPrefix, label):
    return formatBundles(bundles, linkPrefix, label, formatCompanyBundle)

def formatContactBundles(bundles, linkPrefix, label):
    return formatBundles(bundles, linkPrefix, label, formatContactBundle)

def formatMarketingTargetBundles(bundles, linkPrefix):
    return formatBundles(bundles, linkPrefix, 'Marketing Targets', formatMarketingTargetBundle)

def formatPatentBundles(bundles, linkPrefix):
    return formatBundles(bundles, linkPrefix, 'Patents', formatPatentBundle)

def formatTechnologyBundles(bundles, linkPrefix):
    return formatBundles(bundles, linkPrefix, 'Technologies', formatTechnologyBundle)


# Methods

def formatAgreement(agreement):
    return templateAgreementLabel % agreement

def formatCompany(company):
    return templateCompanyLabel % company

def formatContact(contact):
    return templateContactLabel % contact

def formatDocument(document):
    return templateDocumentLabel % document

def formatContactName(contact):
    return (contact['firstName'] + ' ' + contact['middleName']).strip() + ' ' + contact['lastName']

def formatFold(link, label, substance, showSubstance=False):
    # If we have no substance,
    if not substance:
        return ''
    # Make content
    contentPart1 = "<div id=%s class=tabOFF onclick=toggle(this) onmouseover=on(this) onmouseout=off(this)>\n%s\n</div>" % (link, label)
    contentPart2 = "<div id=%s_ class=%s>%s</div>" % (link, 'show' if showSubstance else 'hide', substance)
    # Return
    return contentPart1 + contentPart2

def formatLabel(items, label):
    return '%s (%s)' % (label, len(items))

def formatMarketingProject(marketingProject):
    return templateMarketingProjectLabel % marketingProject

def formatMarketingTarget(marketingTarget):
    return templateMarketingTargetLabel % marketingTarget

def formatPatent(patent):
    if 'leadInventor' not in patent:
        patent['leadInventor'] = ''
    return templatePatentLabel % patent

def formatPayable(payable):
    return templatePayableLabel % payable

def formatPayableDetails(payableDetails, linkPrefix):
    # Set
    link = linkPrefix + '-payableDetails'
    contents = []
    # For each payableDetail,
    for payableDetail in payableDetails:
        # Eat
        contents.append(formatFold(link + '%(payableDetailID)s' % payableDetail, templateTable % templatePayableDetailLabel % payableDetail, templatePayableDetailSubstance % payableDetail))
    # Return
    return formatFold(link, formatLabel(payableDetails, 'Expenses'), '\n'.join(contents))

def formatPayment(payment):
    return templatePaymentLabel % payment

def formatReceivable(receivable):
    return templateReceivableLabel % receivable

def formatRemark(remark):
    if 'remarkSummary' not in remark:
        remark['remarkSummary'] = summarizeRemark(remark['remark'])
    return templateRemarkLabel % remark

def formatRemarks(remarks, linkPrefix):
    # Set
    link = linkPrefix + '-remarks'
    contents = []
    # For each remark,
    for remark in remarks:
        # Eat
        contents.append(formatFold(link + '%(remarkID)s' % remark, templateTable % formatRemark(remark), "<div class=remark>%s</div>" % remark['remark'].replace('\n', '<br>')))
    # Return
    return formatFold(link, formatLabel(remarks, 'Remarks'), '\n'.join(contents))

def formatTechnology(technology):
    if 'leadInventor' not in technology:
        technology['leadInventor'] = ''
    return templateTechnologyLabel % technology

def summarizeRemark(text):
    # Try to extract the subject
    for line in text.splitlines():
        match = patternSubject.match(line)
        if match: 
            return match.group(1).strip()
    # Collapse whitespace
    text = patternWhitespace.sub(' ', text.strip())
    # Get first 100 characters
    return text[:100]

def unCamel(text):
    return re.sub(r'I D$', r'ID', re.sub(r'([A-Z])', r' \1', text).strip().title())


# Constants

patternSubject = re.compile('Subject: (.*)')
patternWhitespace = re.compile('\s+')
formatPackByLabel = {
    'Agreements': (formatAgreementBundle, formatAgreement),
    'Companies': (formatCompanyBundle, formatCompany),
    'Contacts': (formatContactBundle, formatContact),
    'Documents': (formatDocumentBundle, formatDocument),
    'Marketing Projects': (formatMarketingProjectBundle, formatMarketingProject),
    'Marketing Targets': (formatMarketingTargetBundle, formatMarketingTarget),
    'Patents': (formatPatentBundle, formatPatent),
    'Remarks': (formatRemarkBundle, formatRemark),
    'Technologies': (formatTechnologyBundle, formatTechnology),
}
labelByTable = {
    'agrmnts': 'Agreements',
    'company': 'Companies',
    'contacts': 'Contacts',
    'document': 'Documents',
    'mktprj': 'Marketing Projects',
    'mkttgt': 'Marketing Targets',
    'patents': 'Patents',
    'remarks': 'Remarks',
    'technol': 'Technologies',
}
