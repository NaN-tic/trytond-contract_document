import unittest

from proteus import Model, Wizard
from trytond.modules.contract_document.tests.tools import setup
from trytond.tests.test_tryton import drop_db


class TestContactType(unittest.TestCase):

    def setUp(self):
        drop_db()
        super().setUp()

    def tearDown(self):
        drop_db()
        super().tearDown()

    def test(self):
        vars = setup()

        Contract = Model.get('contract')

        contract = Contract(vars.contract.id)
        lessor_contact = contract.lessor_document_contacts.new()
        lessor_contact.name = 'Lessor Contact'
        lessee_contact = contract.lessee_document_contacts.new()
        lessee_contact.name = 'Lessee Contact'
        contract.save()

        contract = Contract(vars.contract.id)
        self.assertEqual(len(contract.lessor_document_contacts), 1)
        self.assertEqual(len(contract.lessee_document_contacts), 1)
        self.assertEqual(contract.lessor_document_contacts[0].type, 'lessor')
        self.assertEqual(contract.lessor_document_contacts[0].name,
            'Lessor Contact')
        self.assertEqual(contract.lessee_document_contacts[0].type, 'lessee')
        self.assertEqual(contract.lessee_document_contacts[0].name,
            'Lessee Contact')

        wizard = Wizard('contract.generate', [contract])
        wizard.form.lessor_company = vars.lessor
        wizard.form.lessor_contact = vars.lessor
        wizard.form.contract_base = vars.contract_base

        wizard_lessor = wizard.form.lessor_document_contacts.new()
        wizard_lessor.name = 'Wizard Lessor'
        wizard_lessee = wizard.form.lessee_document_contacts.new()
        wizard_lessee.name = 'Wizard Lessee'

        wizard.execute('generate')

        contract = Contract(vars.contract.id)
        lessor_names = [c.name for c in contract.lessor_document_contacts]
        lessee_names = [c.name for c in contract.lessee_document_contacts]
        self.assertEqual(len(contract.lessor_document_contacts), 2)
        self.assertEqual(len(contract.lessee_document_contacts), 2)
        self.assertIn('Lessor Contact', lessor_names)
        self.assertIn('Wizard Lessor', lessor_names)
        self.assertIn('Lessee Contact', lessee_names)
        self.assertIn('Wizard Lessee', lessee_names)
        self.assertEqual(
            {c.type for c in contract.lessor_document_contacts}, {'lessor'})
        self.assertEqual(
            {c.type for c in contract.lessee_document_contacts}, {'lessee'})
