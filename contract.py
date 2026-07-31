# The COPYRIGHT file at the top level of this repository contains
# the full copyright notices and license terms.
from decimal import ROUND_HALF_UP
from decimal import Decimal
from io import BytesIO
from datetime import datetime
from datetime import date
import os
import shutil
import subprocess
import tempfile
import zipfile
from xml.sax.saxutils import escape

from jinja2 import ChainableUndefined, Environment
from jinja2.exceptions import TemplateError
from sql.aggregate import Max
from trytond.i18n import gettext
from trytond.model import (
    ModelSQL, ModelView, fields, sequence_ordered, tree)
from trytond.pool import Pool, PoolMeta
from trytond.pyson import Bool, Eval
from trytond.tools import grouped_slice, reduce_ids
from trytond.transaction import Transaction
from trytond.wizard import Button, StateTransition, StateView, Wizard
from trytond.exceptions import UserError, UserWarning

from .tools import (SafeDict, TemplateRecord, markdown_to_paragraphs,
    safe_text, template_value)


def get_jinja_environment():
    return Environment(
        autoescape=False,
        undefined=ChainableUndefined,
        trim_blocks=False,
        lstrip_blocks=False)


def get_jinja_error(location, exc, kind):
    return UserError(
        'The contract document contains an invalid Jinja %s in %s: %s: %s'
        % (kind, location, exc.__class__.__name__, exc))


def get_jinja_syntax_error(location, exc):
    return get_jinja_error(location, exc, 'syntax')


def build_jinja_location(section, field, name=None):
    location = section
    if name:
        location = '%s "%s"' % (location, name)
    return '%s / %s' % (location, field)


def validate_jinja_source(source, location):
    source = (source or '').replace('\r\n', '\n')
    try:
        return get_jinja_environment().from_string(source)
    except TemplateError as exc:
        raise get_jinja_syntax_error(location, exc)


def render_jinja_source(source, context, location):
    template = validate_jinja_source(source, location)
    try:
        return template.render(**context)
    except (TemplateError, TypeError, ValueError) as exc:
        raise get_jinja_error(location, exc, 'expression')


def get_sync_group_key(attribute_set_id, attributes):
    return (
        attribute_set_id,
        tuple(sorted((attributes or {}).items())),
        )


def get_unique_records(records):
    unique = {}
    for record in records:
        if record and record.id:
            unique[record.id] = record
    return list(unique.values())


class JinjaValidationMixin:
    __slots__ = ()

    _jinja_section_name = None

    @classmethod
    def __setup__(cls):
        super().__setup__()
        cls._buttons.update({
                'validate_jinja': {},
                })

    @classmethod
    @ModelView.button
    def validate_jinja(cls, records):
        for record in records:
            record._validate_jinja_fields()

    def _get_jinja_section_name(self):
        return self._jinja_section_name or self.__doc__ or self.__name__

    def _get_jinja_validation_fields(self):
        return [('content', 'content')]

    def _get_jinja_record_name(self):
        return safe_text(getattr(self, 'rec_name', None)).strip() or None

    def _validate_jinja_fields(self):
        name = self._get_jinja_record_name()
        section = self._get_jinja_section_name()
        for field_name, field_label in self._get_jinja_validation_fields():
            validate_jinja_source(
                getattr(self, field_name, None),
                build_jinja_location(section, field_label, name=name))


class Contract(metaclass=PoolMeta):
    __name__ = 'contract'

    cadastre = fields.Char('Cadastre')
    grace_period_months = fields.Integer('Grace Period Months')
    index_price = fields.Numeric('Index Price', digits=(16, 2))
    previous_contract_updated_price = fields.Numeric(
        'Previous Contract Updated Price', digits=(16, 2))
    document_attribute_set = fields.Many2One('asset.attribute.set',
        'Document Attribute Set')
    document_attributes = fields.Dict('asset.attribute',
        'Document Attributes', domain=[
            ('sets', '=', Eval('document_attribute_set', -1)),
            ], states={
            'readonly': Bool(Eval('state')) & (Eval('state') != 'draft'),
            }, depends=['document_attribute_set', 'state'])
    lessor_document_contacts = fields.One2Many('contract.document.contact',
        'contract', 'Lessor Persons', domain=[
            ('type', '=', 'lessor'),
            ], filter=[
            ('type', '=', 'lessor'),
            ], context={
            'default_type': 'lessor',
            })
    lessee_document_contacts = fields.One2Many('contract.document.contact',
        'contract', 'Lessee Persons', domain=[
            ('type', '=', 'lessee'),
            ], filter=[
            ('type', '=', 'lessee'),
            ], context={
            'default_type': 'lessee',
            })
    contract_end_date = fields.Function(fields.Date('Contract End Date'),
        'get_contract_end_date')

    @classmethod
    def get_contract_end_date(cls, contracts, name):
        pool = Pool()
        ContractLine = pool.get('contract.line')
        cursor = Transaction().connection.cursor()
        line = ContractLine.__table__()
        result = {}
        contract_ids = [c.id for c in contracts]
        for sub_ids in grouped_slice(contract_ids):
            where = reduce_ids(line.contract, sub_ids)
            cursor.execute(*line.select(line.contract,
                    Max(line.contract_end_date),
                    where=where,
                    group_by=line.contract))
            for contract_id, value in cursor.fetchall():
                if isinstance(value, str):
                    value = date(*[int(x) for x in value.split('-')])
                result[contract_id] = value
        return {c.id: result.get(c.id) for c in contracts}

    @classmethod
    def __setup__(cls):
        super().__setup__()
        cls._buttons.update({
                'sync_document_attributes_to_assets': {},
                })

    @classmethod
    def create(cls, vlist):
        contracts = super().create(vlist)
        if Transaction().context.get('skip_contract_document_asset_sync'):
            return contracts
        to_sync = []
        for contract, values in zip(contracts, vlist):
            if cls._has_document_attribute_changes(values):
                to_sync.append(contract)
        if to_sync:
            cls._sync_document_attributes_to_assets(to_sync)
        return contracts

    @classmethod
    def write(cls, *args):
        if Transaction().context.get('skip_contract_document_asset_sync'):
            super().write(*args)
            return
        Warning = Pool().get('res.user.warning')
        to_sync = []
        actions = iter(args)
        for contracts, values in zip(actions, actions):
            if not cls._has_document_attribute_changes(values):
                continue
            changed_contracts = [c for c in contracts
                if cls._document_attributes_changed(c, values)]
            to_sync.extend(changed_contracts)
            multiple = [c for c in contracts
                if c in changed_contracts
                and len(c.get_related_assets()) > 1]
            if not multiple:
                continue
            key = Warning.format('contract_document_contract_asset_sync',
                multiple)
            if Warning.check(key):
                raise UserWarning(key, gettext(
                        'contract_document.msg_contract_multiple_assets_warning',
                        contracts=', '.join(c.rec_name for c in multiple)))
        super().write(*args)
        if to_sync:
            cls._sync_document_attributes_to_assets(
                cls.browse(list({c.id for c in to_sync})))

    @classmethod
    def _has_document_attribute_changes(cls, values):
        return any(name in values
            for name in ('document_attribute_set', 'document_attributes'))

    @classmethod
    def _document_attributes_changed(cls, contract, values):
        if 'document_attribute_set' in values:
            current_set = (contract.document_attribute_set.id
                if contract.document_attribute_set else None)
            if values['document_attribute_set'] != current_set:
                return True
        if 'document_attributes' in values:
            if dict(values['document_attributes'] or {}) != dict(
                    contract.document_attributes or {}):
                return True
        return False

    def get_related_assets(self):
        return get_unique_records(
            line.asset for line in self.lines if getattr(line, 'asset', None))

    def set_document_attributes_from_asset_if_empty(self, asset):
        if not asset:
            return
        if (not self.document_attribute_set
                and getattr(asset, 'attribute_set', None)):
            self.document_attribute_set = asset.attribute_set
        if (not self.document_attributes
                and getattr(asset, 'attributes', None)):
            self.document_attributes = dict(asset.attributes)

    @fields.depends('lines', 'document_attribute_set', 'document_attributes')
    def on_change_lines(self):
        if self.document_attribute_set or self.document_attributes:
            return
        for line in self.lines or []:
            if getattr(line, 'asset', None):
                self.set_document_attributes_from_asset_if_empty(line.asset)
                if self.document_attribute_set or self.document_attributes:
                    break

    @classmethod
    def _get_asset_sync_values(cls, contract):
        return {
            'attribute_set': (contract.document_attribute_set.id
                if contract.document_attribute_set else None),
            'attributes': dict(contract.document_attributes or {}),
            }

    @classmethod
    def _sync_document_attributes_to_assets(cls, contracts, force=False):
        Asset = Pool().get('asset')
        grouped_assets = {}
        for contract in contracts:
            assets = contract.get_related_assets()
            if not assets:
                continue
            if len(assets) > 1 and not force:
                continue
            values = cls._get_asset_sync_values(contract)
            key = get_sync_group_key(values['attribute_set'],
                values['attributes'])
            grouped_assets.setdefault(key, {
                    'values': values,
                    'assets': set(),
                    })
            grouped_assets[key]['assets'].update(asset.id for asset in assets)
        if not grouped_assets:
            return
        with Transaction().set_context(skip_asset_attribute_sync=True):
            for data in grouped_assets.values():
                Asset.write(list(Asset.browse(list(data['assets']))),
                    data['values'])

    @classmethod
    @ModelView.button
    def sync_document_attributes_to_assets(cls, contracts):
        Warning = Pool().get('res.user.warning')
        multiple = [c for c in contracts if len(c.get_related_assets()) > 1]
        if multiple:
            key = Warning.format('contract_document_sync_assets',
                multiple)
            if Warning.check(key):
                raise UserWarning(key, gettext(
                        'contract_document.msg_sync_multiple_assets_warning',
                        contracts=', '.join(c.rec_name for c in multiple)))
        cls._sync_document_attributes_to_assets(contracts, force=True)


class ContractContactRole(ModelSQL, ModelView):
    'Contract Contact Role'
    __name__ = 'contract.document.contact.role'

    name = fields.Char('Name', required=True, translate=True)

    def get_rec_name(self, name):
        return self.name


class ContractLineType(ModelSQL, ModelView):
    'Contract Line Type'
    __name__ = 'contract.line.type'

    name = fields.Char('Name', required=True, translate=True)

    def get_rec_name(self, name):
        return self.name


class ContractLine(metaclass=PoolMeta):
    __name__ = 'contract.line'

    line_type = fields.Many2One('contract.line.type', 'Line Type')
    contract_end_date = fields.Date('Contract End Date',
        help='Date when the contract ends. This does not mean billing stops, '
        'as contracts may be automatically renewed unless terminated by either party.')

    @classmethod
    def __setup__(cls):
        super().__setup__()
        cls.end_date.help = (
            'Date when billing ends for this contract line. '
            'This is different from the contract end date, '
            'which indicates when the contract itself ends.')


class ContractContact(sequence_ordered(), ModelSQL, ModelView):
    'Contract Document Contact'
    __name__ = 'contract.document.contact'

    contract = fields.Many2One('contract', 'Contract', required=True,
        ondelete='CASCADE')
    type = fields.Selection([
            ('lessor', 'Lessor'),
            ('lessee', 'Lessee'),
            ], 'Type', required=True)
    name = fields.Char('Full Name', required=True)
    identifier = fields.Char('Identifier')
    address = fields.Text('Address')
    mobile = fields.Char('Mobile')
    email = fields.Char('Email')
    acting_as = fields.Many2One('contract.document.contact.role',
        'Acting As')

    @staticmethod
    def default_type():
        return Transaction().context.get('default_type')

    @classmethod
    def create(cls, vlist):
        default_type = Transaction().context.get('default_type')
        if default_type:
            vlist = [dict(values, type=values.get('type') or default_type)
                for values in vlist]
        return super().create(vlist)

    def get_rec_name(self, name):
        return self.name


class ContractClause(
        JinjaValidationMixin,
        tree(separator=' / '), sequence_ordered(), ModelSQL, ModelView):
    'Contract Clause'
    __name__ = 'contract.document.clause'
    _jinja_section_name = 'Contract Clause'

    name = fields.Char('Name', required=True)
    title = fields.Char('Title', translate=True)
    parent = fields.Many2One('contract.document.clause', 'Parent',
        ondelete='CASCADE', domain=[
            ('id', '!=', Eval('id', -1)),
            ('parent', 'not child_of', [Eval('id', -1)]),
            ], depends=['id'])
    children = fields.One2Many('contract.document.clause', 'parent',
        'Children')
    content = fields.Text('Content', translate=True,
        help='Supports Jinja2 placeholders like {{ contract_number }} '
        'or {{ asset_name }}.')
    active = fields.Boolean('Active')

    @staticmethod
    def default_active():
        return True

    def get_rec_name(self, name):
        return self.name


class ContractBase(JinjaValidationMixin, ModelSQL, ModelView):
    'Contract Base'
    __name__ = 'contract.document.base'
    _jinja_section_name = 'Contract Base'

    name = fields.Char('Name', required=True)
    contract_title = fields.Char('Contract Title', translate=True,
        help='Supports Jinja2 placeholders like {{ contract_number }} '
        'or {{ asset_name }}.')
    parties = fields.One2Many('contract.document.base.party', 'base',
        'Parties')
    appearances = fields.One2Many('contract.document.base.appearance', 'base',
        'Appearances')
    statements = fields.One2Many('contract.document.base.statement', 'base',
        'Statements')
    clauses = fields.One2Many('contract.document.base.clause', 'base',
        'Clauses')

    def _get_jinja_validation_fields(self):
        return [('contract_title', 'contract title')]


class ContractBaseParty(sequence_ordered(), ModelSQL, ModelView):
    'Contract Base Party Text'
    __name__ = 'contract.document.base.party'

    base = fields.Many2One('contract.document.base', 'Base', required=True,
        ondelete='CASCADE')
    party = fields.Many2One('contract.document.party', 'Party Text',
        required=True, ondelete='RESTRICT')


class ContractBaseAppearance(sequence_ordered(), ModelSQL, ModelView):
    'Contract Base Appearance Text'
    __name__ = 'contract.document.base.appearance'

    base = fields.Many2One('contract.document.base', 'Base', required=True,
        ondelete='CASCADE')
    appearance = fields.Many2One('contract.document.appearance',
        'Appearance Text', required=True, ondelete='RESTRICT')


class ContractBaseStatement(sequence_ordered(), ModelSQL, ModelView):
    'Contract Base Statement'
    __name__ = 'contract.document.base.statement'

    base = fields.Many2One('contract.document.base', 'Base', required=True,
        ondelete='CASCADE')
    statement = fields.Many2One('contract.document.manifest', 'Statement',
        required=True, ondelete='RESTRICT')


class ContractBaseClause(sequence_ordered(), ModelSQL, ModelView):
    'Contract Base Clause'
    __name__ = 'contract.document.base.clause'

    base = fields.Many2One('contract.document.base', 'Base', required=True,
        ondelete='CASCADE')
    clause = fields.Many2One('contract.document.clause', 'Clause',
        required=True, ondelete='RESTRICT', domain=[
            ('parent', '=', None),
            ])


class ContractManifest(
        JinjaValidationMixin, sequence_ordered(), ModelSQL, ModelView):
    'Contract Manifest'
    __name__ = 'contract.document.manifest'
    _jinja_section_name = 'Contract Manifest'

    name = fields.Char('Name', required=True)
    title = fields.Char('Title', translate=True)
    content = fields.Text('Content', translate=True,
        help='Supports Jinja2 placeholders like {{ lessor_company }} '
        'or {{ asset_address }}.')
    active = fields.Boolean('Active')

    @staticmethod
    def default_active():
        return True


class ContractParty(
        JinjaValidationMixin, sequence_ordered(), ModelSQL, ModelView):
    'Contract Party Text'
    __name__ = 'contract.document.party'
    _jinja_section_name = 'Contract Party Text'

    name = fields.Char('Name', required=True)
    title = fields.Char('Title', translate=True)
    content = fields.Text('Content', translate=True,
        help='Supports Jinja2 placeholders like {{ lessor_company }} '
        'or {{ asset_address }}.')
    active = fields.Boolean('Active')

    @staticmethod
    def default_active():
        return True


class ContractAppearance(
        JinjaValidationMixin, sequence_ordered(), ModelSQL, ModelView):
    'Contract Appearance Text'
    __name__ = 'contract.document.appearance'
    _jinja_section_name = 'Contract Appearance Text'

    name = fields.Char('Name', required=True, translate=True)
    title = fields.Char('Title', translate=True)
    content = fields.Text('Content', translate=True,
        help='Supports Jinja2 placeholders like {{ lessor_company }} '
        'or {{ asset_address }}.')
    active = fields.Boolean('Active')

    @staticmethod
    def default_active():
        return True


class ContractGenerateClause(ModelView):
    'Contract Generate Clause'
    __name__ = 'contract.generate.start.clause'

    sequence = fields.Integer('Sequence')
    clause = fields.Many2One('contract.document.clause', 'Clause',
        required=True, domain=[
            ('parent', '=', None),
            ])
    title = fields.Char('Title')

    @fields.depends('clause', '_parent_clause.title')
    def on_change_clause(self):
        try:
            title = self.title
        except AttributeError:
            title = None
        if self.clause and not title:
            self.title = self.clause.title


class ContractGenerateStatement(ModelView):
    'Contract Generate Statement'
    __name__ = 'contract.generate.start.statement'

    sequence = fields.Integer('Sequence')
    statement = fields.Many2One('contract.document.manifest', 'Statement')
    title = fields.Char('Title')
    content = fields.Text('Content')

    @fields.depends('statement')
    def on_change_statement(self):
        if not self.statement:
            return
        try:
            title = self.title
        except AttributeError:
            title = None
        try:
            content = self.content
        except AttributeError:
            content = None
        if not title:
            self.title = self.statement.title
        if not content:
            self.content = self.statement.content


class ContractGenerateParty(ModelView):
    'Contract Generate Party Text'
    __name__ = 'contract.generate.start.party'

    sequence = fields.Integer('Sequence')
    party = fields.Many2One('contract.document.party', 'Party Text')
    title = fields.Char('Title')
    content = fields.Text('Content')

    @fields.depends('party')
    def on_change_party(self):
        if not self.party:
            return
        try:
            title = self.title
        except AttributeError:
            title = None
        try:
            content = self.content
        except AttributeError:
            content = None
        if not title:
            self.title = self.party.title
        if not content:
            self.content = self.party.content


class ContractGenerateAppearance(ModelView):
    'Contract Generate Appearance Text'
    __name__ = 'contract.generate.start.appearance'

    sequence = fields.Integer('Sequence')
    appearance = fields.Many2One('contract.document.appearance',
        'Appearance Text')
    title = fields.Char('Title')
    content = fields.Text('Content')

    @fields.depends('appearance')
    def on_change_appearance(self):
        if not self.appearance:
            return
        try:
            title = self.title
        except AttributeError:
            title = None
        try:
            content = self.content
        except AttributeError:
            content = None
        if not title:
            self.title = self.appearance.title
        if not content:
            self.content = self.appearance.content


class ContractGenerateAttachment(ModelView):
    'Contract Generate Attachment'
    __name__ = 'contract.generate.start.attachment'

    sequence = fields.Integer('Sequence')
    name = fields.Char('Name')
    data = fields.Binary('Data', filename='name')


class ContractGenerateContact(ModelView):
    'Contract Generate Contact'
    __name__ = 'contract.generate.start.contact'

    sequence = fields.Integer('Sequence')
    type = fields.Selection([
            ('lessor', 'Lessor'),
            ('lessee', 'Lessee'),
            ], 'Type', required=True)
    name = fields.Char('Full Name', required=True)
    identifier = fields.Char('Identifier')
    address = fields.Text('Address')
    mobile = fields.Char('Mobile')
    email = fields.Char('Email')
    acting_as = fields.Many2One('contract.document.contact.role',
        'Acting As')

    @staticmethod
    def default_type():
        return Transaction().context.get('default_type')


class ContractGenerateStart(ModelView):
    'Generate Contract'
    __name__ = 'contract.generate.start'

    company = fields.Many2One('company.company', 'Company', readonly=True)
    origin = fields.Reference('Origin', selection='get_origin', readonly=True)
    contract_base = fields.Many2One('contract.document.base', 'Contract Base')
    clauses = fields.One2Many('contract.generate.start.clause', None,
        'Clauses')
    contract_title = fields.Char('Contract Title')
    parties_title = fields.Char('Parties Title')
    parties = fields.One2Many('contract.generate.start.party', None,
        'Parties')
    appearances_title = fields.Char('Appearances Title')
    appearances = fields.One2Many('contract.generate.start.appearance', None,
        'Appearances')
    statements_title = fields.Char('Statements Title')
    statements = fields.One2Many('contract.generate.start.statement', None,
        'Statements')
    clauses_title = fields.Char('Clauses Title')
    lessor_company = fields.Many2One('party.party', 'Lessor Company',
        context={
            'company': Eval('company', -1),
            }, depends=['company'])
    lessor_contact = fields.Many2One('party.party', 'Lessor Contact',
        context={
            'company': Eval('company', -1),
            }, depends=['company'])
    lessor_document_contacts = fields.One2Many(
        'contract.generate.start.contact', None, 'Lessor Persons',
        domain=[('type', '=', 'lessor')],
        filter=[('type', '=', 'lessor')], context={
            'default_type': 'lessor',
            })
    payment_type = fields.Many2One('account.payment.type', 'Payment Type',
        domain=[
            ('kind', 'in', ['both', 'receivable']),
            ])
    bank_account = fields.Many2One('bank.account', 'Bank Account')
    lessee_company = fields.Many2One('party.party', 'Lessee Company',
        context={
            'company': Eval('company', -1),
            }, depends=['company'])
    lessee_contacts = fields.Many2One('party.party', 'Lessee Contact',
        context={
            'company': Eval('company', -1),
            }, depends=['company'])
    lessee_document_contacts = fields.One2Many(
        'contract.generate.start.contact', None, 'Lessee Persons',
        domain=[('type', '=', 'lessee')],
        filter=[('type', '=', 'lessee')], context={
            'default_type': 'lessee',
            })
    start_date = fields.Date('Start Date')
    end_date = fields.Date('End Date')
    contract_end_date = fields.Date('Contract End Date',
        help='Date when the contract ends. This does not mean billing stops, '
        'as contracts may be automatically renewed unless terminated by either party.')
    contract_years = fields.Function(fields.Numeric('Contract Years',
            digits=(16, 2)), 'on_change_with_contract_years')
    asset = fields.Many2One('asset', 'Asset', context={
            'company': Eval('company', -1),
            }, depends=['company'])
    deposit = fields.Numeric('Deposit', digits=(16, 2))
    guarantee_amount = fields.Numeric('Guarantee Amount', digits=(16, 2))
    amount = fields.Numeric('Amount', digits=(16, 2))
    cadastre = fields.Char('Cadastre')
    home_assessment = fields.Char('Home Assessment')
    energy_certificate = fields.Char('Energy Certificate')
    attribute_set = fields.Many2One('asset.attribute.set', 'Attribute Set')
    attributes = fields.Dict('asset.attribute', 'Attributes', domain=[
            ('sets', '=', Eval('attribute_set', -1)),
            ], depends=['attribute_set'])
    attachments = fields.One2Many(
        'contract.generate.start.attachment', None, 'Attachments')
    origin_attachments = fields.Many2Many('ir.attachment', None, None,
        'Contract Attachments', domain=[
            ('resource', '=', Eval('origin', -1)),
            ], depends=['origin'])
    sign_digitally = fields.Boolean('Sign Digitally')
    certificate = fields.Many2One('certificate', 'Certificate', context={
            'company': Eval('company', -1),
            }, depends=['company', 'sign_digitally'], states={
            'invisible': ~Eval('sign_digitally', False),
            'required': Bool(Eval('sign_digitally', False)),
            })

    @classmethod
    def get_origin(cls):
        Model = Pool().get('ir.model')
        return [(None, '')] + [('contract', Model.get_name('contract'))]

    @staticmethod
    def default_parties_title():
        return gettext('contract_document.msg_default_parties_title')

    @staticmethod
    def default_appearances_title():
        return gettext('contract_document.msg_default_appearances_title')

    @staticmethod
    def default_statements_title():
        return gettext('contract_document.msg_default_statements_title')

    @staticmethod
    def default_clauses_title():
        return gettext('contract_document.msg_default_clauses_title')

    @fields.depends('contract_base', 'contract_title', 'parties',
        'appearances', 'statements', 'clauses')
    def on_change_contract_base(self):
        pool = Pool()
        ClauseLine = pool.get('contract.generate.start.clause')
        PartyLine = pool.get('contract.generate.start.party')
        AppearanceLine = pool.get('contract.generate.start.appearance')
        StatementLine = pool.get('contract.generate.start.statement')
        if not self.contract_base:
            return
        if self.contract_base.contract_title:
            self.contract_title = self.contract_base.contract_title
        if self.contract_base.parties:
            parties = []
            for index, line in enumerate(sorted(self.contract_base.parties,
                        key=lambda l: ((l.sequence is None), l.sequence or 0,
                            l.id or 0)), start=1):
                party_line = PartyLine()
                party_line.sequence = index
                party_line.party = line.party
                party_line.title = line.party.title
                party_line.content = line.party.content
                parties.append(party_line)
            self.parties = parties
        if self.contract_base.appearances:
            appearances = []
            for index, line in enumerate(sorted(self.contract_base.appearances,
                        key=lambda l: ((l.sequence is None), l.sequence or 0,
                            l.id or 0)), start=1):
                appearance_line = AppearanceLine()
                appearance_line.sequence = index
                appearance_line.appearance = line.appearance
                appearance_line.title = line.appearance.title
                appearance_line.content = line.appearance.content
                appearances.append(appearance_line)
            self.appearances = appearances
        if self.contract_base.statements:
            statements = []
            for index, line in enumerate(sorted(self.contract_base.statements,
                        key=lambda l: ((l.sequence is None), l.sequence or 0,
                            l.id or 0)), start=1):
                statement_line = StatementLine()
                statement_line.sequence = index
                statement_line.statement = line.statement
                statement_line.title = line.statement.title
                statement_line.content = line.statement.content
                statements.append(statement_line)
            self.statements = statements
        clauses = []
        lines = sorted(self.contract_base.clauses,
            key=lambda l: ((l.sequence is None), l.sequence or 0, l.id or 0))
        for index, line in enumerate(lines, start=1):
            clause_line = ClauseLine()
            clause_line.sequence = index
            clause_line.clause = line.clause
            clause_line.title = line.clause.title
            clauses.append(clause_line)
        self.clauses = clauses

    @fields.depends('asset', 'cadastre', 'home_assessment',
        'energy_certificate', 'attribute_set', 'attributes')
    def on_change_asset(self):
        if not self.asset:
            return
        if not self.cadastre:
            self.cadastre = (getattr(self.asset, 'land_register', None)
                or getattr(self.asset, 'home_assessment', None)
                or '')
        if not self.home_assessment:
            self.home_assessment = getattr(self.asset, 'home_assessment', '')
        if not self.energy_certificate:
            self.energy_certificate = getattr(self.asset,
                'energy_certificate', '')
        if not self.attribute_set and getattr(self.asset, 'attribute_set', None):
            self.attribute_set = self.asset.attribute_set
        if not self.attributes and getattr(self.asset, 'attributes', None):
            self.attributes = dict(self.asset.attributes)

    @fields.depends('start_date', 'contract_end_date')
    def on_change_with_contract_years(self, name=None):
        if not self.start_date or not self.contract_end_date:
            return None
        days = (self.contract_end_date - self.start_date).days
        if days <= 0:
            return Decimal('0.00')
        years = Decimal(days) / Decimal('365')
        return years.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)


class ContractGenerateWizard(Wizard):
    'Generate Contract'
    __name__ = 'contract.generate'

    start = StateView('contract.generate.start',
        'contract_document.contract_generate_start_view_form', [
            Button('Cancel', 'end', 'tryton-cancel'),
            Button('Generate', 'generate', 'tryton-ok', default=True),
            ])
    generate = StateTransition()

    @classmethod
    def __setup__(cls):
        super().__setup__()

    def default_start(self, fields_):
        contract = self._get_contract()
        asset = self._get_default_asset(contract)
        lessor = contract.company.party if contract.company else None
        lessee = contract.party
        attributes = dict(contract.document_attributes or {})
        attribute_set = contract.document_attribute_set
        if asset:
            if not attribute_set:
                attribute_set = getattr(asset, 'attribute_set', None)
            if not attributes and getattr(asset, 'attributes', None):
                attributes = dict(asset.attributes)
        contract_title = self._default_contract_title(contract)
        return {
            'company': contract.company.id if contract.company else None,
            'origin': str(contract),
            'contract_title': contract_title,
            'parties': self._default_parties(),
            'appearances': self._default_appearances(),
            'lessor_company': lessor.id if lessor else None,
            'lessor_contact': lessor.id if lessor else None,
            'lessor_document_contacts': self._default_document_contacts(
                contract, 'lessor'),
            'payment_type': (contract.payment_type.id
                if getattr(contract, 'payment_type', None) else None),
            'bank_account': (contract.bank_account.id
                if getattr(contract, 'bank_account', None) else None),
            'lessee_company': lessee.id if lessee else None,
            'lessee_contacts': lessee.id if lessee else None,
            'lessee_document_contacts': self._default_document_contacts(
                contract, 'lessee'),
            'start_date': contract.start_date,
            'end_date': getattr(contract, 'end_date', None),
            'contract_end_date': getattr(contract, 'contract_end_date', None),
            'asset': asset.id if asset else None,
            'deposit': getattr(contract, 'deposit', None),
            'guarantee_amount': getattr(contract, 'guarantee_amount', None),
            'amount': self._get_contract_amount(contract),
            'cadastre': contract.cadastre or self._get_asset_cadastre(asset),
            'home_assessment': getattr(asset, 'home_assessment', None)
            if asset else None,
            'energy_certificate': getattr(asset, 'energy_certificate', None)
            if asset else None,
            'attribute_set': attribute_set.id if attribute_set else None,
            'attributes': attributes,
            'origin_attachments': [],
            'statements': self._default_statements(),
            }

    def transition_generate(self):
        pool = Pool()
        Attachment = pool.get('ir.attachment')
        contract = self._get_contract()

        self._write_back_contract(contract)
        context = self._get_render_context(contract)
        docx_data = self._build_docx(context)
        filename = self._get_output_name(contract, 'docx')
        data = docx_data
        if self.start.sign_digitally:
            pdf_data = self._convert_docx_to_pdf(docx_data, filename)
            data = self._sign_pdf(pdf_data)
            filename = self._get_output_name(contract, 'pdf')

        attachment = Attachment(
            name=filename,
            type='data',
            data=data,
            resource=contract)
        attachment.save()

        if hasattr(contract, 'contract'):
            contract.contract = attachment
            contract.save()
        return 'end'

    def _get_contract(self):
        pool = Pool()
        Contract = pool.get('contract')
        active_ids = Transaction().context.get('active_ids', [])
        if not active_ids:
            raise UserError(gettext(
                'contract_document.msg_missing_active_contract'))
        if len(active_ids) != 1:
            raise UserError(gettext('contract_document.msg_single_contract'))
        return Contract(active_ids[0])

    def _get_contract_attachments(self, contract):
        Attachment = Pool().get('ir.attachment')
        return [a.id for a in Attachment.search([
                    ('resource', '=', str(contract)),
                ])]

    def _get_default_asset(self, contract):
        for line in contract.lines:
            if getattr(line, 'asset', None):
                return line.asset

    def _get_contract_amount(self, contract):
        reference_date = contract.start_date
        amount = Decimal('0.0')
        for line in contract.lines:
            if reference_date:
                if line.start_date and line.start_date > reference_date:
                    continue
                if line.end_date and line.end_date <= reference_date:
                    continue
            quantity = line.quantity
            if quantity is None:
                quantity = 1
            unit_price = Decimal(str(line.unit_price or 0))
            amount += unit_price * Decimal(str(quantity))
        return amount

    def _get_asset_cadastre(self, asset):
        if not asset:
            return ''
        return (getattr(asset, 'land_register', None)
            or getattr(asset, 'home_assessment', None)
            or '')

    def _default_contract_title(self, contract):
        number = contract.number or contract.reference or ''
        return gettext('contract_document.msg_default_contract_title',
            contract=safe_text(number).strip())

    def _default_parties(self):
        return self._default_section_lines(
            'contract.document.party',
            'contract.generate.start.party',
            'party')

    def _default_appearances(self):
        return self._default_section_lines(
            'contract.document.appearance',
            'contract.generate.start.appearance',
            'appearance')

    def _default_statements(self):
        return self._default_section_lines(
            'contract.document.manifest',
            'contract.generate.start.statement',
            'statement')

    def _default_section_lines(self, model_name, line_model_name, field_name):
        pool = Pool()
        Record = pool.get(model_name)
        Line = pool.get(line_model_name)
        lines = []
        for index, record in enumerate(Record.search([],
                    order=[('sequence', 'ASC'), ('id', 'ASC')]), start=1):
            line = Line()
            line.sequence = index
            setattr(line, field_name, record)
            line.title = record.title
            line.content = record.content
            lines.append(line)
        return lines

    def _default_document_contacts(self, contract, contact_type):
        pool = Pool()
        ContactLine = pool.get('contract.generate.start.contact')
        field_name = '%s_document_contacts' % contact_type
        lines = []
        for index, record in enumerate(getattr(contract, field_name, []),
                start=1):
            line = ContactLine()
            line.sequence = index
            line.type = contact_type
            line.name = record.name
            line.identifier = record.identifier
            line.address = record.address
            line.mobile = record.mobile
            line.email = record.email
            line.acting_as = record.acting_as
            lines.append(line)
        return lines

    def _write_back_contract(self, contract):
        pool = Pool()
        Contact = pool.get('contract.document.contact')

        def line_value(line, name, default=None):
            return getattr(line, name, default)

        contract.cadastre = self.start.cadastre
        contract.document_attribute_set = self.start.attribute_set
        contract.document_attributes = self.start.attributes or {}
        contract.save()
        to_delete = Contact.search([
                ('contract', '=', contract.id),
                ('type', 'in', ['lessor', 'lessee']),
                ])
        if to_delete:
            Contact.delete(to_delete)
        to_create = []
        for contact_type, lines in (
                ('lessor', self.start.lessor_document_contacts or []),
                ('lessee', self.start.lessee_document_contacts or [])):
            for index, line in enumerate(lines, start=1):
                if not line.name:
                    continue
                to_create.append({
                        'contract': contract.id,
                        'sequence': index,
                        'type': contact_type,
                        'name': line.name,
                        'identifier': line_value(line, 'identifier'),
                        'address': line_value(line, 'address'),
                        'mobile': line_value(line, 'mobile'),
                        'email': line_value(line, 'email'),
                        'acting_as': (line_value(line, 'acting_as').id
                            if line_value(line, 'acting_as') else None),
                        })
        if to_create:
            Contact.create(to_create)

    def _get_render_context(self, contract):
        asset = self.start.asset or self._get_default_asset(contract)
        lessor_company = self.start.lessor_company
        lessor_contact = self.start.lessor_contact
        payment_type = self.start.payment_type
        bank_account = self.start.bank_account
        lessee_company = self.start.lessee_company
        lessee_contact = self.start.lessee_contacts
        lessee_contacts = [lessee_contact] if lessee_contact else []
        lessor_document_contacts = self._wrap_document_contacts(
            self.start.lessor_document_contacts or [])
        lessee_document_contacts = self._wrap_document_contacts(
            self.start.lessee_document_contacts or [])
        addresses = []
        if asset and getattr(asset, 'current_address', None):
            addresses.append(asset.current_address.rec_name)
        if asset and getattr(asset, 'addresses', None):
            for address in asset.addresses:
                if address.address and address.address.rec_name not in addresses:
                    addresses.append(address.address.rec_name)

        lessee_contact_names = [
            p.rec_name for p in lessee_contacts if p and p.rec_name]
        wrapped_lessee_contacts = [
            TemplateRecord(p) for p in lessee_contacts if p]
        attributes = dict(self.start.attributes or {})
        attribute_set = self.start.attribute_set
        certificate = self.start.certificate
        attachment_names = [
            attachment.name for attachment in self.start.origin_attachments
            if attachment and attachment.name]
        context = SafeDict({
                'today': datetime.now().strftime('%d/%m/%Y'),
                'contract_number': safe_text(contract.number),
                'contract_reference': safe_text(contract.reference),
                'contract': TemplateRecord(contract),
                'contract_lines': [TemplateRecord(line)
                    for line in contract.lines],
                'contract_party': TemplateRecord(contract.party)
                if contract.party else '',
                'contract_party_name': safe_text(contract.party.rec_name
                    if contract.party else ''),
                'company': TemplateRecord(contract.company)
                if contract.company else '',
                'company_name': safe_text(contract.company.rec_name
                    if contract.company else ''),
                'lessor_company': TemplateRecord(lessor_company)
                if lessor_company else '',
                'lessor_company_name': safe_text(lessor_company.rec_name
                    if lessor_company else ''),
                'lessor_contact': TemplateRecord(lessor_contact)
                if lessor_contact else '',
                'lessor_contact_name': safe_text(lessor_contact.rec_name
                    if lessor_contact else ''),
                'lessor_document_contacts': lessor_document_contacts,
                'lessor_document_contacts_text': self._contacts_text(
                    lessor_document_contacts),
                'payment_type': TemplateRecord(payment_type)
                if payment_type else '',
                'payment_type_name': safe_text(payment_type.rec_name
                    if payment_type else ''),
                'bank_account': TemplateRecord(bank_account)
                if bank_account else '',
                'bank_account_name': safe_text(bank_account.rec_name
                    if bank_account else ''),
                'lessee_company': TemplateRecord(lessee_company)
                if lessee_company else '',
                'lessee_company_name': safe_text(lessee_company.rec_name
                    if lessee_company else ''),
                'lessee_contact': TemplateRecord(lessee_contact)
                if lessee_contact else '',
                'lessee_contact_name': safe_text(lessee_contact.rec_name
                    if lessee_contact else ''),
                'lessee_contacts': wrapped_lessee_contacts,
                'lessee_contacts_text': ', '.join(lessee_contact_names),
                'lessee_document_contacts': lessee_document_contacts,
                'lessee_document_contacts_text': self._contacts_text(
                    lessee_document_contacts),
                'start_date': self.start.start_date,
                'start_date_text': safe_text(self.start.start_date),
                'end_date': self.start.end_date,
                'end_date_text': safe_text(self.start.end_date),
                'contract_end_date': self.start.contract_end_date,
                'contract_end_date_text': safe_text(self.start.contract_end_date),
                'contract_years': template_value(self.start.contract_years),
                'contract_years_text': safe_text(self.start.contract_years),
                'asset': TemplateRecord(asset) if asset else '',
                'asset_name': safe_text(asset.rec_name if asset else ''),
                'asset_address': self._get_asset_address(asset) or '; '.join(addresses),
                'deposit': safe_text(self.start.deposit),
                'deposit_value': template_value(self.start.deposit),
                'guarantee_amount': safe_text(self.start.guarantee_amount),
                'guarantee_amount_value': template_value(
                    self.start.guarantee_amount),
                'amount': safe_text(self.start.amount),
                'amount_value': template_value(self.start.amount),
                'cadastre': safe_text(self.start.cadastre),
                'home_assessment': safe_text(self.start.home_assessment),
                'energy_certificate': safe_text(
                    self.start.energy_certificate),
                'attribute_set': TemplateRecord(attribute_set)
                if attribute_set else '',
                'attribute_set_name': safe_text(attribute_set.rec_name
                    if attribute_set else ''),
                'sign_digitally': bool(self.start.sign_digitally),
                'certificate': TemplateRecord(certificate)
                if certificate else '',
                'certificate_name': safe_text(certificate.rec_name
                    if certificate else ''),
                'first_invoice_date': safe_text(contract.first_invoice_date),
                'currency': safe_text(contract.currency.rec_name
                    if contract.currency else ''),
                'attributes': attributes,
                'origin_attachments': attachment_names,
                'origin_attachments_text': '\n'.join(attachment_names),
                })
        for key, value in attributes.items():
            context['attribute_%s' % key] = safe_text(value)
        context['attributes_block'] = '\n'.join(
            '%s: %s' % (key, value)
            for key, value in sorted(attributes.items()))
        context['attachments_block'] = context['origin_attachments_text']
        return context

    def _wrap_document_contacts(self, contacts):
        return [TemplateRecord(contact) for contact in contacts if contact]

    def _contacts_text(self, contacts):
        return ', '.join(contact.name for contact in contacts if contact.name)

    def _get_asset_address(self, asset):
        if not asset:
            return ''
        parts = []
        if getattr(asset, 'road_type', None):
            parts.append(safe_text(asset.road_type))
        if getattr(asset, 'street', None):
            parts.append(safe_text(asset.street))
        number_parts = []
        if getattr(asset, 'number_type', None):
            number_parts.append(safe_text(asset.number_type))
        if getattr(asset, 'number', None):
            number_parts.append(safe_text(asset.number))
        if getattr(asset, 'number_qualifier', None):
            number_parts.append(safe_text(asset.number_qualifier))
        if number_parts:
            parts.append(' '.join(number_parts))
        for label, value in (
                ('Bloc', getattr(asset, 'block', None)),
                ('Portal', getattr(asset, 'doorway', None)),
                ('Escala', getattr(asset, 'stair', None)),
                ('Planta', getattr(asset, 'floor', None)),
                ('Porta', getattr(asset, 'door', None)),
                ):
            if value:
                parts.append('%s %s' % (label, value))
        if getattr(asset, 'complement', None):
            parts.append(safe_text(asset.complement))
        locality = []
        if getattr(asset, 'zip', None):
            locality.append(safe_text(asset.zip))
        if getattr(asset, 'municipality', None):
            locality.append(safe_text(asset.municipality))
        elif getattr(asset, 'city', None):
            locality.append(safe_text(asset.city))
        if locality:
            parts.append(' '.join(locality))
        return ', '.join(x for x in parts if x)

    def _render_text(self, text, context):
        return render_jinja_source(text, context, 'Document / content')

    def _render_template_field(self, text, context, section, field, name=None):
        return render_jinja_source(
            text, context, build_jinja_location(section, field, name=name))

    def _get_related_record_name(self, section_line, field_name, fallback=None):
        record = getattr(section_line, field_name, None)
        if record:
            return safe_text(record.rec_name).strip() or None
        return safe_text(fallback).strip() or None

    def _build_docx(self, context):
        paragraphs = []
        if self.start.contract_title:
            rendered_contract_title = self._render_template_field(
                self.start.contract_title, context, 'Contract', 'title')
            paragraphs.append({
                    'text': rendered_contract_title,
                    'bold': True,
                    'center': True,
                    })
            paragraphs.append({'text': ''})

        self._append_line_section(paragraphs, self.start.parties_title,
            self.start.parties, context, 'Party', 'party')
        self._append_line_section(paragraphs, self.start.appearances_title,
            self.start.appearances, context, 'Appearance', 'appearance')

        rendered_statements = []
        for statement in sorted(self.start.statements,
                key=lambda l: ((l.sequence is None), l.sequence or 0,
                    l.id or 0)):
            name = self._get_related_record_name(
                statement, 'statement', fallback=statement.title)
            rendered_statements.append({
                    'title': statement.title,
                    'content': self._render_template_field(
                        statement.content, context, 'Statement', 'content',
                        name=name),
                    })
        rendered_statements = [s for s in rendered_statements
            if (s['title'] and s['title'].strip())
            or (s['content'] and s['content'].strip())]
        if rendered_statements and self.start.statements_title:
            paragraphs.append({
                    'text': self.start.statements_title,
                    'bold': True,
                    'center': True,
                    })
            for index, statement in enumerate(rendered_statements, start=1):
                if statement['title']:
                    paragraphs.append({
                            'text': '%s. %s' % (index, statement['title']),
                            'bold': True,
                            })
                self._append_markdown(paragraphs, statement['content'])
                paragraphs.append({'text': ''})

        seen = set()
        ordered_clauses = []
        for line in sorted(self.start.clauses,
                key=lambda l: ((l.sequence is None), l.sequence or 0, l.id or 0)):
            if line.clause:
                self._append_clause_tree(line.clause, ordered_clauses, seen)

        rendered_clauses = []
        for clause in ordered_clauses:
            rendered_clauses.append({
                    'title': clause.title,
                    'content': self._render_template_field(
                        clause.content, context, 'Clause', 'content',
                        name=safe_text(clause.rec_name).strip() or None),
                    })
        rendered_clauses = [c for c in rendered_clauses
            if (c['title'] and c['title'].strip())
            or (c['content'] and c['content'].strip())]
        if rendered_clauses and self.start.clauses_title:
            paragraphs.append({'text': ''})
            paragraphs.append({
                    'text': self.start.clauses_title,
                    'bold': True,
                    'center': True,
                    })
            for index, clause in enumerate(rendered_clauses, start=1):
                title = clause['title']
                if title:
                    paragraphs.append({
                            'text': '%s. %s' % (index, title),
                            'bold': True,
                            })
                self._append_markdown(paragraphs, clause['content'])
                paragraphs.append({'text': ''})

        return self._create_docx(paragraphs)

    def _append_line_section(self, paragraphs, title, lines, context,
            section, relation_field):
        rendered_lines = []
        for section_line in sorted(lines,
                key=lambda l: ((l.sequence is None), l.sequence or 0,
                    l.id or 0)):
            name = self._get_related_record_name(
                section_line, relation_field, fallback=section_line.title)
            rendered_lines.append({
                    'title': section_line.title,
                    'content': self._render_template_field(
                        section_line.content, context, section, 'content',
                        name=name),
                    })
        rendered_lines = [line for line in rendered_lines
            if (line['title'] and line['title'].strip())
            or (line['content'] and line['content'].strip())]
        if not title or not rendered_lines:
            return
        paragraphs.append({
                'text': title,
                'bold': True,
                'center': True,
                })
        for section_line in rendered_lines:
            if section_line['title']:
                paragraphs.append({
                        'text': section_line['title'],
                        'bold': True,
                        })
            self._append_markdown(paragraphs, section_line['content'])
            paragraphs.append({'text': ''})
        paragraphs.append({'text': ''})

    def _append_clause_tree(self, clause, ordered_clauses, seen):
        if clause.id in seen:
            return
        seen.add(clause.id)
        ordered_clauses.append(clause)
        children = sorted(clause.children,
            key=lambda c: ((c.sequence is None), c.sequence or 0, c.id or 0))
        for child in children:
            self._append_clause_tree(child, ordered_clauses, seen)

    def _create_docx(self, paragraphs):
        buffer_ = BytesIO()
        with zipfile.ZipFile(buffer_, 'w', zipfile.ZIP_DEFLATED) as zf:
            zf.writestr('[Content_Types].xml', self._content_types_xml())
            zf.writestr('_rels/.rels', self._root_rels_xml())
            zf.writestr('docProps/core.xml', self._core_xml())
            zf.writestr('docProps/app.xml', self._app_xml())
            zf.writestr('word/document.xml', self._document_xml(paragraphs))
        return buffer_.getvalue()

    def _append_markdown(self, paragraphs, text):
        for paragraph in markdown_to_paragraphs(text):
            paragraphs.append(paragraph)

    def _content_types_xml(self):
        return (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/docProps/app.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
            '<Override PartName="/docProps/core.xml" '
            'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
            '<Override PartName="/word/document.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            '</Types>')

    def _root_rels_xml(self):
        return (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="word/document.xml"/>'
            '<Relationship Id="rId2" '
            'Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" '
            'Target="docProps/core.xml"/>'
            '<Relationship Id="rId3" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" '
            'Target="docProps/app.xml"/>'
            '</Relationships>')

    def _core_xml(self):
        return (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<cp:coreProperties '
            'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" '
            'xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<dc:title>Contract document</dc:title>'
            '<dc:creator>Tryton</dc:creator>'
            '</cp:coreProperties>')

    def _app_xml(self):
        return (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            '<Application>Tryton</Application>'
            '</Properties>')

    def _document_xml(self, paragraphs):
        body = ''.join(self._paragraph_xml(p) for p in paragraphs)
        section = (
            '<w:sectPr>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
            'w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
            '</w:sectPr>')
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
            'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
            'xmlns:o="urn:schemas-microsoft-com:office:office" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
            'xmlns:v="urn:schemas-microsoft-com:vml" '
            'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
            'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
            'xmlns:w10="urn:schemas-microsoft-com:office:word" '
            'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
            'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
            'mc:Ignorable="w14 w15">'
            '<w:body>%s%s</w:body></w:document>' % (body, section))

    def _paragraph_xml(self, paragraph):
        ppr = []
        if paragraph.get('center'):
            ppr.append('<w:jc w:val="center"/>')
        ppr.append('<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>')
        ppr_xml = '<w:pPr>%s</w:pPr>' % ''.join(ppr) if ppr else ''
        runs = paragraph.get('runs')
        if runs is None:
            runs = [{
                    'text': paragraph.get('text', ''),
                    'bold': paragraph.get('bold', False),
                    'italic': paragraph.get('italic', False),
                    }]
        if paragraph.get('bullet'):
            runs = [{'text': '• '}] + runs
        return '<w:p>%s%s</w:p>' % (
            ppr_xml, ''.join(self._run_xml(run) for run in runs))

    def _run_xml(self, run):
        if run.get('break'):
            return '<w:r><w:br/></w:r>'
        text = escape(run.get('text', ''))
        rpr = []
        if run.get('bold'):
            rpr.append('<w:b/>')
        if run.get('italic'):
            rpr.append('<w:i/>')
        rpr_xml = '<w:rPr>%s</w:rPr>' % ''.join(rpr) if rpr else ''
        return '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (
            rpr_xml, text)

    def _convert_docx_to_pdf(self, docx_data, filename):
        soffice = shutil.which('soffice') or shutil.which('libreoffice')
        if not soffice:
            raise UserError(gettext('contract_document.msg_missing_soffice'))
        with tempfile.TemporaryDirectory() as directory:
            input_path = os.path.join(directory, filename)
            with open(input_path, 'wb') as handler:
                handler.write(docx_data)
            subprocess.run([
                    soffice,
                    '--headless',
                    '--convert-to', 'pdf',
                    '--outdir', directory,
                    input_path,
                    ], check=True, stdout=subprocess.PIPE,
                stderr=subprocess.PIPE)
            pdf_path = os.path.splitext(input_path)[0] + '.pdf'
            with open(pdf_path, 'rb') as handler:
                return handler.read()

    def _sign_pdf(self, pdf_data):
        if not self.start.certificate:
            return pdf_data
        try:
            from pyhanko.sign import signers
            from pyhanko.pdf_utils.incremental_writer import (
                IncrementalPdfFileWriter)
        except ImportError as exc:
            raise UserError(gettext(
                'contract_document.msg_missing_pyhanko')) from exc

        with tempfile.TemporaryDirectory() as directory:
            source = os.path.join(directory, 'source.pdf')
            signed = os.path.join(directory, 'signed.pdf')
            with open(source, 'wb') as handler:
                handler.write(pdf_data)
            with self.start.certificate.tmp_ssl_credentials() as (crt, key):
                signer = signers.SimpleSigner.load(
                    key_file=key, cert_file=crt, key_passphrase=None)
                metadata = signers.PdfSignatureMetadata(field_name='Signature1')
                pdf_signer = signers.PdfSigner(metadata, signer=signer)
                with open(source, 'rb') as infile, open(signed, 'wb') as output:
                    writer = IncrementalPdfFileWriter(infile)
                    pdf_signer.sign_pdf(writer, output=output)
            with open(signed, 'rb') as handler:
                return handler.read()

    def _get_output_name(self, contract, extension):
        base = contract.number or contract.reference or str(contract.id)
        base = ''.join(c if c.isalnum() or c in ('-', '_') else '_'
            for c in safe_text(base)).strip('_') or 'contract'
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        suffix = 'signed' if extension == 'pdf' else 'draft'
        return '%s_%s_%s.%s' % (base, timestamp, suffix, extension)
