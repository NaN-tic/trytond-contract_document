import io
import unittest
import zipfile
from decimal import Decimal

from proteus import Model, Wizard
from trytond.modules.contract_document.tests.tools import setup
from trytond.tests.test_tryton import drop_db


class TestGenerateContract(unittest.TestCase):

    def setUp(self):
        drop_db()
        super().setUp()

    def tearDown(self):
        drop_db()
        super().tearDown()

    def test(self):
        vars = setup()

        Attachment = Model.get('ir.attachment')

        wizard = Wizard('contract.generate', [vars.contract])
        self.assertEqual(wizard.form.payment_type, vars.payment_type)
        self.assertEqual(wizard.form.bank_account, vars.bank_account)
        self.assertEqual(wizard.form.asset, vars.asset)
        self.assertEqual(wizard.form.start_date.isoformat(), '2015-01-01')
        self.assertEqual(wizard.form.end_date.isoformat(), '2016-01-01')
        self.assertEqual(wizard.form.contract_years, Decimal('1.00'))

        wizard.form.lessor_company = vars.lessor
        wizard.form.lessor_contact = vars.lessor
        wizard.form.contract_base = vars.contract_base

        self.assertEqual(len(wizard.form.clauses), 1)
        self.assertEqual(wizard.form.clauses[0].title, 'Clause Title')
        self.assertEqual(wizard.form.parties[0].title, 'Party Title')
        self.assertEqual(wizard.form.appearances[0].title, 'Appearance Title')
        self.assertEqual(wizard.form.statements[0].title, 'Statement Title')

        wizard.execute('generate')

        attachments = Attachment.find([('resource', '=', vars.contract)])
        self.assertEqual(len(attachments), 1)
        attachment, = attachments
        self.assertTrue(attachment.name.endswith('.docx'))

        with zipfile.ZipFile(io.BytesIO(bytes(attachment.data)), 'r') as docx:
            document_xml = docx.read('word/document.xml').decode('utf-8')

        self.assertIn('PARTIES', document_xml)
        self.assertIn('Party Title', document_xml)
        self.assertIn('Lessor Company', document_xml)
        self.assertIn('Appearance Title', document_xml)
        self.assertIn('Receivable', document_xml)
        self.assertIn('ES76 2077 0024 0031 0257 5766', document_xml)
        self.assertIn('Statement Title', document_xml)
        self.assertIn('1.00', document_xml)
        self.assertIn('Clause Title', document_xml)
        self.assertIn('Apartment 1A', document_xml)
        self.assertIn('750', document_xml)
