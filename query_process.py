'Methods for finding and reporting information'
# Import system modules
import os
import re
import csv
import collections
import StringIO
# Import custom modules
import config
import message_format
import model_database


class Processor(object):
    'Query processor'

    # Constructor

    def __init__(self, processor):
        'Initialize'
        self.processor = processor
        self.database = processor.database

    # Process

    def process(self):
        'Parse and run query'
        # Test for query
        match = re.match(r'\?(.*)', self.processor.message['subject'])
        # If we have a query,
        if match:
            # Extract terms
            termString = match.group(1).strip()
            # Process query or return instructions
            messagePack = self.processQuery(termString) if termString else self.processHelp()
            # Send results
            self.processor.queueWrite(*messagePack)
            # Move to folderQuery
            self.processor.queueMove(config.folderQuery)
            # Return
            return True
        
    def processQuery(self, termString):
        'Run and assemble query'
        # Prepare
        terms = set(parseTermString(termString))
        resultPacksByLabel = self.getResultPacksByLabel(terms)
        # Initialize
        summaries = []
        attachmentPaths = []
        # Make temporary folder
        temporaryFolder = self.processor.makeTemporaryFolder()
        # For each label,
        for label in sorted(resultPacksByLabel):
            # Load
            resultPacks = resultPacksByLabel[label]
            # Append summary
            summaries.append('<tr><td valign=top>%s</td><td>%s</td></tr>' % (label, ', '.join('%s (%d)' % (' '.join(termsIncluded), len(bundles)) for termsIncluded, bundles in resultPacks)))
            # Append attachment HTML
            attachmentPath = os.path.join(temporaryFolder, label + '.html')
            attachmentPaths.append(attachmentPath)
            attachmentContent = message_format.formatQueryResults(label, resultPacks, terms)
            with open(attachmentPath, 'wt') as attachmentFile:
                attachmentFile.write(attachmentContent)
            # Append attachment CSV
            attachmentPath = os.path.join(temporaryFolder, label + '.csv')
            attachmentPaths.append(attachmentPath)
            csvWriter = csv.writer(open(attachmentPath, 'wb'))
            for termsIncluded, bundles in resultPacks:
                csvWriter.writerow(['Matching terms: ' + ' '.join(termsIncluded)])
                keys = sorted(bundles[0][0].keys())
                csvWriter.writerow([message_format.unCamel(x) for x in keys])
                for bundle in bundles:
                    csvWriter.writerow([bundle[0][x] for x in keys])
                csvWriter.writerow([''])
        # Return
        return message_format.templateQuery % '<br>'.join(summaries), attachmentPaths

    def processHelp(self):
        'Show instructions'
        return message_format.formatHelp(), None

    # Get

    def getResultPacksByLabel(self, terms):
        'Return resultPacks arranged by result type'
        # Initialize
        resultPacksByLabel = collections.defaultdict(list)
        objectIDsByTermsByLabel = self.getObjectIDsByTermsByLabel(terms)
        # For each label,
        for label, objectIDsByTerms in objectIDsByTermsByLabel.iteritems():
            # Prepare method
            getBundle = self.database.getBundleByLabel[label]
            # Sort results in order of most number of included terms first
            for termsIncluded in sorted(objectIDsByTerms, key=lambda x: (len(x), x), reverse=True):
                # Get bundles
                objectIDs = objectIDsByTerms[termsIncluded]
                bundles = [getBundle(x) for x in objectIDs]
                # Append
                resultPacksByLabel[label].append((termsIncluded, bundles))
        # Return
        return resultPacksByLabel

    def getObjectIDsByTermsByLabel(self, terms):
        'Return ordered objectIDs arranged by included terms for each result type'
        # Initialize
        termOrdersByIDByLabel = collections.defaultdict(dict)
        # For each term,
        for term in terms:
            # Tally companies
            companyIDs = self.database.findCompanyIDs(term)
            companyIDString = model_database.makeIDString(companyIDs)
            aggregateTermOrdersByID(term, termOrdersByIDByLabel['Companies'], companyIDs)
            # Tally contacts
            contactIDs = self.database.findContactIDs(term)
            contactIDString = model_database.makeIDString(contactIDs)
            aggregateTermOrdersByID(term, termOrdersByIDByLabel['Contacts'], contactIDs)
            # Tally remarks
            remarkIDs = self.database.findRemarkIDs(term)
            remarkIDString = model_database.makeIDString(remarkIDs)
            aggregateTermOrdersByID(term, termOrdersByIDByLabel['Remarks'], remarkIDs)
            # Tally technologies
            technologyIDs = self.database.findTechnologyIDs(term, contactIDString)
            technologyIDString = model_database.makeIDString(technologyIDs)
            aggregateTermOrdersByID(term, termOrdersByIDByLabel['Technologies'], technologyIDs)
            # Tally patents
            patentIDs = self.database.findPatentIDs(term, companyIDString, contactIDString, technologyIDString)
            patentIDString = model_database.makeIDString(patentIDs)
            aggregateTermOrdersByID(term, termOrdersByIDByLabel['Patents'], patentIDs)
            # Tally agreements
            agreementIDs = self.database.findAgreementIDs(term, companyIDString, contactIDString, technologyIDString)
            agreementIDString = model_database.makeIDString(agreementIDs)
            aggregateTermOrdersByID(term, termOrdersByIDByLabel['Agreements'], agreementIDs)
            # Tally marketing projects
            marketingProjectIDs = self.database.findMarketingProjectIDs(term, companyIDString, contactIDString, patentIDString, technologyIDString)
            marketingProjectIDString = model_database.makeIDString(marketingProjectIDs)
            aggregateTermOrdersByID(term, termOrdersByIDByLabel['Marketing Projects'], marketingProjectIDs)
            # Tally marketing targets
            marketingTargetIDs = self.database.findMarketingTargetIDs(term, companyIDString, contactIDString, marketingProjectIDString, patentIDString, technologyIDString)
            marketingTargetIDString = model_database.makeIDString(marketingTargetIDs)
            aggregateTermOrdersByID(term, termOrdersByIDByLabel['Marketing Targets'], marketingTargetIDs)
            # Tally documents
            documentIDs = self.database.findDocumentIDs(term, agreementIDString, companyIDString, contactIDString, marketingProjectIDString, marketingTargetIDString, patentIDString, remarkIDString, technologyIDString)
            documentIDString = model_database.makeIDString(documentIDs)
            aggregateTermOrdersByID(term, termOrdersByIDByLabel['Documents'], documentIDs)
        # Return
        return dict((label, gatherObjectIDsByTerms(termOrdersByID)) for label, termOrdersByID in termOrdersByIDByLabel.iteritems())


# Helpers

def parseTermString(termString):
    return csv.reader(StringIO.StringIO(termString.replace(',', '')), delimiter=' ').next()

def aggregateTermOrdersByID(term, termOrdersByID, objectIDs):
    'Store term and order information for each objectID'
    # For each objectID,
    for objectID in objectIDs:
        # Get term and term's order because otherwise we will lose order information
        termOrder = term, objectIDs.index(objectID)
        # If it doesn't exist,
        if objectID not in termOrdersByID:
            termOrdersByID[objectID] = []
        # Append
        termOrdersByID[objectID].append(termOrder)

def gatherObjectIDsByTerms(termOrdersByID):
    'Bin objectIDs by included terms'
    # Initialize
    orderIDsByTerms = collections.defaultdict(list)
    # For each termOrder,
    for objectID, termOrders in termOrdersByID.iteritems():
        # Get terms
        terms = tuple(term for term, order in termOrders)
        # Use order of the first term
        order = termOrders[0][1]
        # Append
        orderIDsByTerms[terms].append((order, objectID))
    # Sort by order and bin by included terms
    return dict((terms, [objectID for order, objectID in sorted(orderIDs)]) for terms, orderIDs in orderIDsByTerms.iteritems())
