import datetime
from decimal import Decimal
from types import SimpleNamespace

from proteus import Model
from trytond.modules.account.tests.tools import (
    create_chart, create_fiscalyear, get_accounts)
from trytond.modules.account_invoice.tests.tools import (
    create_payment_term, set_fiscalyear_invoice_sequences)
from trytond.modules.company.tests.tools import create_company, get_company
from trytond.tests.tools import activate_modules


def setup():
    activate_modules('contract_document')

    today = datetime.date(2015, 1, 1)

    _ = create_company()
    company = get_company()

    fiscalyear = set_fiscalyear_invoice_sequences(
        create_fiscalyear(company, today))
    fiscalyear.click('create_period')

    _ = create_chart(company)
    accounts = get_accounts(company)
    revenue = accounts['revenue']
    expense = accounts['expense']

    payment_term = create_payment_term()
    payment_term.save()

    PaymentType = Model.get('account.payment.type')
    payment_type = PaymentType(
        name='Receivable',
        kind='receivable',
        account_bank='party')
    payment_type.save()

    Party = Model.get('party.party')
    bank_party = Party(name='Main Bank')
    bank_party.save()

    Bank = Model.get('bank')
    bank = Bank()
    bank.party = bank_party
    bank.save()

    BankAccount = Model.get('bank.account')
    BankNumber = Model.get('bank.account.number')
    bank_account = BankAccount()
    bank_account.bank = bank
    bank_account.numbers.append(BankNumber(number='ES7620770024003102575766',
            type='iban'))
    bank_account.save()

    lessor = Party(name='Lessor Company')
    lessor.save()
    customer = Party(name='Customer Company')
    customer.customer_payment_term = payment_term
    customer.customer_payment_type = payment_type
    customer.bank_accounts.append(bank_account)
    customer.receivable_bank_account = bank_account
    customer.save()

    ProductCategory = Model.get('product.category')
    account_category = ProductCategory(name='Account Category')
    account_category.accounting = True
    account_category.account_expense = expense
    account_category.account_revenue = revenue
    account_category.save()

    ProductUom = Model.get('product.uom')
    unit, = ProductUom.find([('name', '=', 'Unit')])

    ProductTemplate = Model.get('product.template')
    asset_template = ProductTemplate()
    asset_template.name = 'Apartment Asset'
    asset_template.default_uom = unit
    asset_template.type = 'assets'
    asset_template.list_price = Decimal('1000')
    asset_template.cost_price_method = 'fixed'
    asset_template.account_category = account_category
    asset_template.save()
    asset_product, = asset_template.products

    service_template = ProductTemplate()
    service_template.name = 'Rent Service'
    service_template.default_uom = unit
    service_template.type = 'service'
    service_template.list_price = Decimal('750')
    service_template.cost_price = Decimal('0')
    service_template.cost_price_method = 'fixed'
    service_template.account_category = account_category
    service_template.save()
    service_product, = service_template.products

    Asset = Model.get('asset')
    AssetAttribute = Model.get('asset.attribute')
    AssetAttributeSet = Model.get('asset.attribute.set')

    document_attribute = AssetAttribute()
    document_attribute.name = 'registry'
    document_attribute.string = 'Registry'
    document_attribute.type_ = 'char'
    document_attribute.save()

    document_attribute_floor = AssetAttribute()
    document_attribute_floor.name = 'floor'
    document_attribute_floor.string = 'Floor'
    document_attribute_floor.type_ = 'char'
    document_attribute_floor.save()

    attribute_set = AssetAttributeSet(name='Document Attributes')
    attribute_set.attributes.append(document_attribute)
    attribute_set.attributes.append(document_attribute_floor)
    attribute_set.save()

    asset = Asset()
    asset.name = 'Apartment 1A'
    asset.product = asset_product
    asset.attribute_set = attribute_set
    asset.attributes = {'registry': 'ASSET-1'}
    asset.save()

    second_asset = Asset()
    second_asset.name = 'Apartment 1B'
    second_asset.product = asset_product
    second_asset.attribute_set = attribute_set
    second_asset.attributes = {'floor': '1B'}
    second_asset.save()

    Service = Model.get('contract.service')
    service = Service(name='Monthly Rent', product=service_product)
    service.save()

    Sequence = Model.get('ir.sequence')
    sequence_contract, = Sequence.find([('name', '=', 'Contract')])
    Journal = Model.get('account.journal')
    journal, = Journal.find([('type', '=', 'revenue')])
    ContractConfig = Model.get('contract.configuration')
    contract_config = ContractConfig(1)
    contract_config.contract_sequence = sequence_contract
    contract_config.journal = journal
    contract_config.save()

    ContractClause = Model.get('contract.document.clause')
    clause = ContractClause(name='internal-clause')
    clause.title = 'Clause Title'
    clause.content = (
        'Asset {{ asset_name }}\n'
        'Amount {{ amount }}\n'
        'Start {{ start_date_text }}')
    clause.save()

    ContractParty = Model.get('contract.document.party')
    party_text = ContractParty(name='internal-party')
    party_text.title = 'Party Title'
    party_text.content = (
        'Lessor {{ lessor_company_name }}\n'
        'Contact {{ lessor_contact_name }}\n'
        'Line Type {{ contract.lines[0].line_type.rec_name }}\n'
        'Service {{ contract.lines[0].service.rec_name }}\n'
        'Quantity {{ contract.lines[0].quantity }}')
    party_text.save()

    ContractAppearance = Model.get('contract.document.appearance')
    appearance = ContractAppearance(name='internal-appearance')
    appearance.title = 'Appearance Title'
    appearance.content = (
        'Payment {{ payment_type_name }}\n'
        'Bank {{ bank_account_name }}')
    appearance.save()

    ContractManifest = Model.get('contract.document.manifest')
    statement = ContractManifest(name='internal-statement')
    statement.title = 'Statement Title'
    statement.content = (
        'Years {{ contract_years_text }}\n'
        'Computed {{ (amount_value + 0.0) if not lessee_company else amount_value }}')
    statement.save()

    ContractBase = Model.get('contract.document.base')
    contract_base = ContractBase(name='Lease Base')
    contract_base.contract_title = 'Base Contract {{ contract_number }}'
    base_party = contract_base.parties.new()
    base_party.party = party_text
    base_appearance = contract_base.appearances.new()
    base_appearance.appearance = appearance
    base_statement = contract_base.statements.new()
    base_statement.statement = statement
    base_clause = contract_base.clauses.new()
    base_clause.clause = clause
    contract_base.save()

    ContactRole = Model.get('contract.document.contact.role')
    contact_role = ContactRole(name='Tenant')
    contact_role.save()
    ContractLineType = Model.get('contract.line.type')
    rent_line_type, = ContractLineType.find([('name', '=', 'Rent')])

    Contract = Model.get('contract')
    contract = Contract()
    contract.company = company
    contract.party = customer
    contract.payment_term = payment_term
    contract.payment_type = payment_type
    contract.bank_account = bank_account
    contract.start_period_date = today
    contract.first_invoice_date = today
    contract.freq = 'monthly'
    contract.interval = 1
    line = contract.lines.new()
    line.service = service
    line.asset = asset
    line.line_type = rent_line_type
    line.unit_price = Decimal('750')
    line.quantity = 2
    line.unit = unit
    line.start_date = today
    line.end_date = datetime.date(2016, 1, 1)
    line.contract_end_date = datetime.date(2016, 1, 1)
    discount_line = contract.lines.new()
    discount_line.service = service
    discount_line.unit_price = Decimal('50')
    discount_line.quantity = -1
    discount_line.unit = unit
    discount_line.start_date = today
    discount_line.end_date = datetime.date(2016, 1, 1)
    discount_line.contract_end_date = datetime.date(2016, 1, 1)
    ended_line = contract.lines.new()
    ended_line.service = service
    ended_line.unit_price = Decimal('100')
    ended_line.quantity = 5
    ended_line.unit = unit
    ended_line.start_date = today
    ended_line.end_date = today
    ended_line.contract_end_date = today
    contract.save()
    contract.click('confirm')

    return SimpleNamespace(
        company=company,
        lessor=lessor,
        customer=customer,
        payment_type=payment_type,
        bank_account=bank_account,
        payment_term=payment_term,
        service=service,
        unit=unit,
        today=today,
        attribute_set=attribute_set,
        document_attribute=document_attribute,
        asset=asset,
        second_asset=second_asset,
        contract=contract,
        contract_base=contract_base,
        contact_role=contact_role)
