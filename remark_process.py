'Methods for detecting and processing messages intended as remarks'
# Import system modules
import collections
import re
import os
# Import custom modules
import config
import message_format
import model_database
import query_process


# Set patterns
pattern_remark = re.compile(r'##([\S]*)\s+(.*)')


class Processor(object):
    'Remark processor'

    # Constructor

    def __init__(self, processor):
        'Initialize'
        self.processor = processor
        self.database = processor.database
        self.processMethodByType = {
             'AGREEMENT': self.processTermsAgreement,
                     'A': self.processTermsAgreement,
               'COMPANY': self.processTermsCompany,
                    'CM': self.processTermsCompany,
               'CONTACT': self.processTermsContact,
                    'CN': self.processTermsContact,
              'MPROJECT': self.processTermsMarketingProject,
                    'MP': self.processTermsMarketingProject,
               'MTARGET': self.processTermsMarketingTarget,
                    'MT': self.processTermsMarketingTarget,
                'PATENT': self.processTermsPatent,
                     'P': self.processTermsPatent,
               # 'PAYABLE': self.processTermsPayable,
                    # 'PY': self.processTermsPayable,
            # 'RECEIVABLE': self.processTermsReceivable,
                    # 'RV': self.processTermsReceivable,
            'TECHNOLOGY': self.processTermsTechnology,
                     'T': self.processTermsTechnology,
        }

    # Process

    def process(self):
        'Parse and add remark'
        # Initialize
        errors = []
        linkIDsByTable = collections.defaultdict(list)
        # For each line of the body,
        for line in self.processor.message['body'].splitlines():
            # Test for remark directions
            match = pattern_remark.match(line)
            # If we have remark directions,
            if match:
                # Convert it to machine directions
                try: 
                    self.processRemark(match, linkIDsByTable)
                # If there was an error,
                except RemarkError, error: 
                    # Save it
                    errors.append(str(error))
        # If there were errors
        if errors:
            # Write response
            self.processor.queueWrite(message_format.formatRemarkFailed(self.processor.message, '<br><br>\n'.join(errors)))
            # Move to folderFailed
            self.processor.queueMove(config.folderFailed)
            # Return
            return True
        # Otherwise, if there were machine directions,
        elif linkIDsByTable: 
            # Execute machine directions
            attachmentNames = self.processLinks(linkIDsByTable)
            # Build confirmation reply
            summaries = []
            for linkTable in sorted(linkIDsByTable):
                # Initialize
                linkIDs = linkIDsByTable[linkTable]
                label = message_format.labelByTable[linkTable]
                getBundle = self.database.getBundleByLabel[label]
                # Unpack
                formatSummary = message_format.formatPackByLabel[label][1]
                # Append summary
                summaries.append('<b>%s</b><br>\n' % label.capitalize() + message_format.templateTable % '</tr><tr>'.join(formatSummary(getBundle(x, simple=True)) for x in linkIDs))
            # Write confirmation reply
            self.processor.queueWrite(message_format.formatRemarkResults(self.processor.message, summaries, attachmentNames))
            # Move to folderRemark
            self.processor.queueMove(config.folderRemark)
            # Return
            return True

    def processRemark(self, match, linkIDsByTable):
        # Extract match
        remarkType, termString = match.groups()
        remarkType = remarkType.upper()
        termStrings = [termString.strip()]
        # If we don't recognize the remark type,
        if remarkType not in self.processMethodByType:
            raise RemarkError('Could not recognize remark type=%s<br>\nPlease use one of the following: %s' % (remarkType, ', '.join(self.processMethodByType.keys())))
        # Execute method
        linkTable, linkIDs = self.processMethodByType[remarkType](termStrings)
        # Add links
        linkIDsByTable[linkTable].extend(linkIDs)

    def processLinks(self, linkIDsByTable):
        # Remove duplicate links
        for linkTable, linkIDs in linkIDsByTable.iteritems():
            linkIDsByTable[linkTable] = list(set(linkIDs))
        # Get attachments
        attachmentPacks = self.processor.extractAttachments()
        attachmentFileNames = [x[0] for x in attachmentPacks]
        attachmentMessage = 'Documents:\n%s\n\n' % '\n'.join(attachmentFileNames) if attachmentFileNames else ''
        # Prepare
        remarkSummary = self.processor.message['subject'] + '\n\n'
        remark = remarkSummary + attachmentMessage + message_format.formatMailOriginal(self.processor.message)
        remarkDate = self.processor.message['when']
        # For each link,
        for linkTable, linkIDs in linkIDsByTable.iteritems():
            for linkID in linkIDs:
                self.database.addRemark(linkTable, linkID, remark, remarkDate)
                for attachmentPack in attachmentPacks:
                    self.database.addDocument(linkTable, linkID, *attachmentPack)
        # Return
        return attachmentFileNames

    def processTermsAgreement(self, terms):
        methods = self.database.findAgreementIDs, self.database.getAgreementBundle, message_format.formatAgreement
        return 'agrmnts', processTerms(terms, methods, 'agreements')

    def processTermsCompany(self, terms):
        methods = self.database.findCompanyIDs, self.database.getCompanyBundle, message_format.formatCompany
        return 'company', processTerms(terms, methods, 'companies')

    def processTermsContact(self, terms):
        methods = self.database.findContactIDs, self.database.getContactBundle, message_format.formatContact
        return 'contacts', processTerms(terms, methods, 'contacts')

    def processTermsMarketingProject(self, terms):
        methods = self.database.findMarketingProjectIDs, self.database.getMarketingProjectBundle, message_format.formatMarketingProject
        return 'mktprj', processTerms(terms, methods, 'marketing projects')

    def processTermsMarketingTarget(self, terms):
        methods = self.database.findMarketingTargetIDs, self.database.getMarketingTargetBundle, message_format.formatMarketingTarget
        return 'mkttgt', processTerms(terms, methods, 'marketing targets')

    def processTermsPatent(self, terms):
        methods = self.database.findPatentIDs, self.database.getPatentBundle, message_format.formatPatent
        return 'patents', processTerms(terms, methods, 'patents')

    # def processTermsPayable(self, terms):
        # methods = self.database.findPayableIDs, self.database.getPayableBundle, message_format.formatPayable
        # return 'payable', processTerms(terms, methods, 'payables')

    # def processTermsReceivable(self, terms):
        # methods = self.database.findReceivableIDs, self.database.getReceivableBundle, message_format.formatReceivable
        # return 'recvable', processTerms(terms, methods, 'receivables')

    def processTermsTechnology(self, terms):
        methods = self.database.findTechnologyIDs, self.database.getTechnologyBundle, message_format.formatTechnology
        return 'technol', processTerms(terms, methods, 'technologies')


def processTerms(termStrings, methods, label):
    # Initialize
    findIDs, getID, formatSummary = methods
    linkIDs, unresolvableTerms, ambiguousPacks = [], [], []
    # For each termString,
    for termString in termStrings:
        # Get resultIDs using AND
        resultIDs = reduce(lambda x, y: x.intersection(y), [set(findIDs(x)) for x in query_process.parseTermString(termString)])
        # If we find nothing,
        if not resultIDs:
            unresolvableTerms.append(termString)
        # If we have more than one match,
        elif len(resultIDs) > 1:
            ambiguousPacks.append((termString, resultIDs))
        # If we have exactly one match, use it
        else: 
            linkIDs.append(resultIDs.pop())
    # If we have unresolvable or ambiguous termStrings,
    if unresolvableTerms or ambiguousPacks:
        # Initialize
        messages = []
        # Display unresolvable termStrings
        if unresolvableTerms:
            message = 'Could not find %s for the following terms<br>' % label
            messages.append(message + '; '.join(unresolvableTerms) + '<br>')
        # Display ambiguous termStrings
        if ambiguousPacks:
            message = 'Multiple ' + label + ' found for "%s"<br>'
            messages.extend(message % ambiguousTerm + message_format.templateTable % '</tr>\n<tr>'.join(formatSummary(getID(x, simple=True)) for x in ambiguousResults) for ambiguousTerm, ambiguousResults in ambiguousPacks)
        # Raise error
        raise RemarkError('<br>\n'.join(messages))
    # Return
    return linkIDs


class RemarkError(Exception):
    pass
