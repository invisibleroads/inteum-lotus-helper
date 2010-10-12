"""
Test whether we can interact with the database properly.

- Load document bundle
- Load remark bundle
- Get document
"""
# Import system modules
import os
import datetime
import unittest
import sys; sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
# Import custom modules
import model_database


class TestModelDatabase(unittest.TestCase):
    
    # Initialize
    
    def setUp(self):
        self.database = model_database.Model()

    # Add remark

    def testAddRemarkAgreement(self): 
        self.addRemark('agrmnts', self.getFirstAgreementID())

    def testAddRemarkCompany(self):
        self.addRemark('company', self.getFirstCompanyID())

    def testAddRemarkContact(self):
        self.addRemark('contacts', self.getFirstContactID())

    def testAddRemarkMarketingProject(self):
        self.addRemark('mktprj', self.getFirstMarketingProjectID())

    def testAddRemarkMarketingTarget(self):
        self.addRemark('mkttgt', self.getFirstMarketingTargetID())

    def testAddRemarkPatent(self): 
        self.addRemark('patents', self.getFirstPatentID())

    def testAddRemarkTechnology(self): 
        self.addRemark('technol', self.getFirstTechnologyID())

    def addRemark(self, linkTable, linkFK):
        # Add a remark
        valueByName = self.database.addRemark(linkTable, linkFK, 'Testing', datetime.datetime.now())
        remarkID = valueByName['primarykey']
        # Check that the remark was added
        remarkDictionaries = self.database.getRemarks(linkTable, linkFK) 
        remarkIDs = [x['remarkID'] for x in remarkDictionaries]
        self.assert_(remarkID in remarkIDs)
        # Delete the remark
        self.database.deleteRemark(remarkID)
        # Check that the remark was deleted
        remarkDictionaries = self.database.getRemarks(linkTable, linkFK) 
        remarkIDs = [x['remarkID'] for x in remarkDictionaries]
        self.assert_(remarkID not in remarkIDs)

    # Get bundle

    def testGetAgreementBundle(self): 
        self.database.getAgreementBundle(self.getFirstAgreementID())

    def testGetCompanyBundle(self): 
        self.database.getCompanyBundle(self.getFirstCompanyID())

    def testGetContactBundle(self): 
        self.database.getContactBundle(self.getFirstContactID())

    def testGetDocumentBundle(self):
        for linkTable in ['agrmnts', 'company', 'contacts', 'mktprj', 'mkttgt', 'patents', 'payable', 'payment', 'recvable', 'technol']:
            self.database.cursor1.execute('SELECT primarykey FROM document WHERE linktable=?', [linkTable])
            linkID = self.database.cursor1.fetchone()[0]
            self.database.getDocumentBundle(linkID)

    # def testGetMarketingProjectBundle(self):
        # self.database.getMarketingProjectBundle(self.getFirstMarketingProjectID())

    # def testGetMarketingTargetBundle(self):
        # self.database.getMarketingTargetBundle(self.getFirstMarketingTargetID())

    def testGetPatentBundle(self): 
        self.database.getPatentBundle(self.getFirstPatentID())

    def testGetTechnologyBundle(self): 
        self.database.getTechnologyBundle(self.getFirstTechnologyID())

    # Get

    def getDocument(self, documentID):
        pass

    def getFirstID(self, tableName):
        sql = 'SELECT primarykey FROM ' + tableName
        self.database.cursor1.execute(sql)
        firstID = self.database.cursor1.fetchone()[0]
        return firstID

    def getFirstAgreementID(self):
        return self.getFirstID('agrmnts')

    def getFirstCompanyID(self):
        return self.getFirstID('company')

    def getFirstContactID(self):
        return self.getFirstID('contacts')

    def getFirstMarketingProjectID(self):
        return self.getFirstID('mktprj')

    def getFirstMarketingTargetID(self):
        return self.getFirstID('mkttgt')

    def getFirstPatentID(self):
        return self.getFirstID('patents')

    def getFirstTechnologyID(self):
        return self.getFirstID('technol')


if __name__ == '__main__':
    unittest.main()
