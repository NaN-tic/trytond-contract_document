import unittest
from decimal import Decimal

from proteus import Model, config
from trytond.exceptions import UserWarning
from trytond.modules.contract_document.tests.tools import setup
from trytond.tests.test_tryton import drop_db


class TestSyncAttributes(unittest.TestCase):

    def setUp(self):
        drop_db()
        super().setUp()

    def tearDown(self):
        drop_db()
        super().tearDown()

    def test(self):
        vars = setup()

        Asset = Model.get('asset')
        Contract = Model.get('contract')
        Warning = Model.get('res.user.warning')

        draft_contract = Contract()
        draft_contract.company = vars.company
        draft_contract.party = vars.customer
        draft_contract.payment_term = vars.payment_term
        draft_contract.payment_type = vars.payment_type
        draft_contract.bank_account = vars.bank_account
        draft_contract.start_period_date = vars.today
        draft_contract.first_invoice_date = vars.today
        draft_contract.freq = 'monthly'
        draft_contract.interval = 1
        line = draft_contract.lines.new()
        line.service = vars.service
        line.asset = vars.asset
        line.unit_price = Decimal('100')
        line.quantity = 1
        line.unit = vars.unit
        line.start_date = vars.today
        draft_contract._on_change(['lines'])
        self.assertEqual(draft_contract.document_attribute_set,
            vars.attribute_set)
        self.assertEqual(dict(draft_contract.document_attributes), {
                'registry': 'ASSET-1',
                })

        contract = Contract(vars.contract.id)
        contract.document_attribute_set = vars.attribute_set
        contract.document_attributes = {'registry': 'CONTRACT-1'}
        contract.save()

        asset = Asset(vars.asset.id)
        self.assertEqual(asset.attribute_set, vars.attribute_set)
        self.assertEqual(dict(asset.attributes), {'registry': 'CONTRACT-1'})

        asset.attributes = {'registry': 'ASSET-UPDATED'}
        asset.save()

        contract = Contract(vars.contract.id)
        self.assertEqual(contract.document_attribute_set, vars.attribute_set)
        self.assertEqual(dict(contract.document_attributes), {
                'registry': 'ASSET-UPDATED',
                })

        multi_contract = Contract()
        multi_contract.company = vars.company
        multi_contract.party = vars.customer
        multi_contract.payment_term = vars.payment_term
        multi_contract.payment_type = vars.payment_type
        multi_contract.bank_account = vars.bank_account
        multi_contract.start_period_date = vars.today
        multi_contract.first_invoice_date = vars.today
        multi_contract.freq = 'monthly'
        multi_contract.interval = 1
        for related_asset in (vars.asset, vars.second_asset):
            line = multi_contract.lines.new()
            line.service = vars.service
            line.asset = related_asset
            line.unit_price = Decimal('100')
            line.quantity = 1
            line.unit = vars.unit
            line.start_date = vars.today
        multi_contract.document_attribute_set = vars.attribute_set
        multi_contract.document_attributes = {'registry': 'MULTI-CONTRACT'}
        multi_contract.save()

        asset = Asset(vars.asset.id)
        second_asset = Asset(vars.second_asset.id)
        self.assertEqual(dict(asset.attributes), {'registry': 'ASSET-UPDATED'})
        self.assertEqual(dict(second_asset.attributes), {'floor': '1B'})

        key = None
        multi_contract.document_attributes = {'registry': 'MULTI-ONLY-CONTRACT'}
        with self.assertRaises(UserWarning):
            try:
                multi_contract.save()
            except UserWarning as warning:
                _, (key, *_) = warning.args
                raise

        Warning.skip(key, True, config.get_config().context)
        multi_contract.save()

        multi_contract = Contract(multi_contract.id)
        asset = Asset(vars.asset.id)
        second_asset = Asset(vars.second_asset.id)
        self.assertEqual(dict(multi_contract.document_attributes), {
                'registry': 'MULTI-ONLY-CONTRACT',
                })
        self.assertEqual(dict(asset.attributes), {'registry': 'ASSET-UPDATED'})
        self.assertEqual(dict(second_asset.attributes), {'floor': '1B'})

        asset.attributes = {'registry': 'MERGED'}
        asset.save()

        multi_contract = Contract(multi_contract.id)
        self.assertEqual(dict(multi_contract.document_attributes), {
                'registry': 'MERGED',
                'floor': '1B',
                })

        multi_contract.document_attributes = {
            'registry': 'MULTI-CONTRACT',
            'floor': '1B',
            }
        multi_contract.save()

        key = None
        with self.assertRaises(UserWarning):
            try:
                multi_contract.click('sync_document_attributes_to_assets')
            except UserWarning as warning:
                _, (key, *_) = warning.args
                raise

        Warning.skip(key, True, config.get_config().context)
        multi_contract.click('sync_document_attributes_to_assets')

        asset = Asset(vars.asset.id)
        second_asset = Asset(vars.second_asset.id)
        self.assertEqual(dict(asset.attributes), {
                'registry': 'MULTI-CONTRACT',
                'floor': '1B',
                })
        self.assertEqual(dict(second_asset.attributes), {
                'registry': 'MULTI-CONTRACT',
                'floor': '1B',
                })
