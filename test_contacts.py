import unittest
from contacts import Store, normalize, eligible, safe_csv

class ContactTests(unittest.TestCase):
    def setUp(self):
        self.store=Store(':memory:')
    def tearDown(self):
        self.store.db.close()
    def test_normalize(self):
        self.assertEqual(normalize('01012345678'),('201012345678','Egypt mobile'))
        self.assertEqual(normalize('٠١٠١٢٣٤٥٦٧٨')[0],'201012345678')
        self.assertEqual(normalize('0221234567')[1],'Landline / other Egypt')
        self.assertEqual(normalize('not a number')[1],'Invalid')
    def test_dedupe_optout(self):
        self.assertEqual(self.store.import_rows([{'phone':'01012345678'},{'phone':'+201012345678'}]),(1,1,0))
        row=self.store.all()[0]
        self.assertFalse(eligible(row))
        self.store.update(row['id'],'Do not contact','Unknown','','')
        self.store.import_rows([{'phone':'01012345678'}])
        self.assertEqual(self.store.get(row['id'])['status'],'Do not contact')
        self.assertFalse(eligible(self.store.get(row['id'])))
    def test_reply_match(self):
        self.store.import_rows([{'phone':'01012345678'}])
        row=self.store.all()[0]
        self.assertIsNone(self.store.match(row['phone']))
        with self.assertRaises(ValueError):self.store.event(row['id'],'Sent','Test')
        self.store.update(row['id'],'New','Confirmed manually','Test permission','')
        self.store.event(row['id'],'Sent','Test')
        self.assertEqual(self.store.match('01012345678')['id'],row['id'])
        self.store.event(row['id'],'Reply','Test reply')
        self.assertEqual(self.store.get(row['id'])['status'],'Replied')
    def test_csv(self):
        self.assertEqual(safe_csv('=1+1'),"'=1+1")
if __name__=='__main__':unittest.main()
