# The COPYRIGHT file at the top level of this repository contains
# the full copyright notices and license terms.
from decimal import Decimal
from io import BytesIO
from datetime import datetime
import os
import shutil
import subprocess
import tempfile
import zipfile
from xml.sax.saxutils import escape

from jinja2 import ChainableUndefined, Environment
from trytond.i18n import gettext
from trytond.model import (
    ModelSQL, ModelView, fields, sequence_ordered, tree)
from trytond.pool import Pool, PoolMeta
from trytond.pyson import Bool, Eval
from trytond.transaction import Transaction
from trytond.wizard import Button, StateTransition, StateView, Wizard
from trytond.exceptions import UserError

from .tools import SafeDict, TemplateRecord, markdown_to_paragraphs, safe_text


class Contract(metaclass=PoolMeta):
    __name__ = 'contract'

    cadastre = fields.Char('Cadastre')
    document_attribute_set = fields.Many2One('asset.attribute.set',
        'Document Attribute Set')
    document_attributes = fields.Dict('asset.attribute',
        'Document Attributes', domain=[
            ('sets', '=', Eval('document_attribute_set', -1)),
            ], states={
            'readonly': Bool(Eval('state')) & (Eval('state') != 'draft'),
            }, depends=['document_attribute_set', 'state'])


class ContractClause(
        tree(separator=' / '), sequence_ordered(), ModelSQL, ModelView):
    'Contract Clause'
    __name__ = 'contract.document.clause'

    name = fields.Char('Name', required=True, translate=True)
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


class ContractBase(ModelSQL, ModelView):
    'Contract Base'
    __name__ = 'contract.document.base'

    name = fields.Char('Name', required=True, translate=True)
    clauses = fields.One2Many('contract.document.base.clause', 'base',
        'Clauses')


class ContractBaseClause(sequence_ordered(), ModelSQL, ModelView):
    'Contract Base Clause'
    __name__ = 'contract.document.base.clause'

    base = fields.Many2One('contract.document.base', 'Base', required=True,
        ondelete='CASCADE')
    clause = fields.Many2One('contract.document.clause', 'Clause',
        required=True, ondelete='RESTRICT', domain=[
            ('parent', '=', None),
            ])


class ContractManifest(sequence_ordered(), ModelSQL, ModelView):
    'Contract Manifest'
    __name__ = 'contract.document.manifest'

    name = fields.Char('Name', required=True, translate=True)
    title = fields.Char('Title', translate=True)
    content = fields.Text('Content', translate=True,
        help='Supports Jinja2 placeholders like {{ lessor_company }} '
        'or {{ asset_address }}.')
    active = fields.Boolean('Active')

    @staticmethod
    def default_active():
        return True


class ContractParty(sequence_ordered(), ModelSQL, ModelView):
    'Contract Party Text'
    __name__ = 'contract.document.party'

    name = fields.Char('Name', required=True, translate=True)
    title = fields.Char('Title', translate=True)
    content = fields.Text('Content', translate=True,
        help='Supports Jinja2 placeholders like {{ lessor_company }} '
        'or {{ asset_address }}.')
    active = fields.Boolean('Active')

    @staticmethod
    def default_active():
        return True


class ContractAppearance(sequence_ordered(), ModelSQL, ModelView):
    'Contract Appearance Text'
    __name__ = 'contract.document.appearance'

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
            self.title = self.clause.title or self.clause.name


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


class ContractGenerateStart(ModelView):
    'Generate Contract'
    __name__ = 'contract.generate.start'

    company = fields.Many2One('company.company', 'Company', readonly=True)
    origin = fields.Reference('Origin', selection='get_origin', readonly=True)
    contract_base = fields.Many2One('contract.document.base', 'Contract Base')
    clauses = fields.One2Many('contract.generate.start.clause', None,
        'Clauses')
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
    lessee_company = fields.Many2One('party.party', 'Lessee Company',
        context={
            'company': Eval('company', -1),
            }, depends=['company'])
    lessee_contacts = fields.Many2Many('party.party', None, None,
        'Lessee Contacts', context={
            'company': Eval('company', -1),
            }, depends=['company'])
    asset = fields.Many2One('asset', 'Asset', context={
            'company': Eval('company', -1),
            }, depends=['company'])
    amount = fields.Numeric('Amount', digits=(16, 2))
    cadastre = fields.Char('Cadastre')
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
        return 'PARTIES'

    @staticmethod
    def default_appearances_title():
        return 'APPEARANCES'

    @staticmethod
    def default_statements_title():
        return 'STATEMENTS'

    @staticmethod
    def default_clauses_title():
        return 'CLAUSES'

    @fields.depends('contract_base', 'clauses')
    def on_change_contract_base(self):
        pool = Pool()
        ClauseLine = pool.get('contract.generate.start.clause')
        if not self.contract_base:
            return
        clauses = []
        lines = sorted(self.contract_base.clauses,
            key=lambda l: ((l.sequence is None), l.sequence or 0, l.id or 0))
        for index, line in enumerate(lines, start=1):
            clause_line = ClauseLine()
            clause_line.sequence = index
            clause_line.clause = line.clause
            clause_line.title = line.clause.title or line.clause.name
            clauses.append(clause_line)
        self.clauses = clauses

    @fields.depends('asset', 'cadastre', 'attribute_set', 'attributes')
    def on_change_asset(self):
        if not self.asset:
            return
        if not self.cadastre:
            self.cadastre = (getattr(self.asset, 'land_register', None)
                or getattr(self.asset, 'home_assessment', None)
                or '')
        if not self.attribute_set and getattr(self.asset, 'attribute_set', None):
            self.attribute_set = self.asset.attribute_set
        if not self.attributes and getattr(self.asset, 'attributes', None):
            self.attributes = dict(self.asset.attributes)


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
        return {
            'company': contract.company.id if contract.company else None,
            'origin': str(contract),
            'parties': self._default_parties(),
            'appearances': self._default_appearances(),
            'lessor_company': lessor.id if lessor else None,
            'lessor_contact': lessor.id if lessor else None,
            'lessee_company': lessee.id if lessee else None,
            'lessee_contacts': [lessee.id] if lessee else [],
            'asset': asset.id if asset else None,
            'amount': self._get_contract_amount(contract),
            'cadastre': contract.cadastre or self._get_asset_cadastre(asset),
            'attribute_set': attribute_set.id if attribute_set else None,
            'attributes': attributes,
            'origin_attachments': [],
            'statements': self._default_statements(),
            }

    def transition_generate(self):
        pool = Pool()
        Attachment = pool.get('ir.attachment')
        contract = self._get_contract()
        if not any(line.clause for line in self.start.clauses or []):
            raise UserError(gettext(
                'contract_document.msg_missing_clauses'))

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
        amount = Decimal('0.0')
        for line in contract.lines:
            unit_price = Decimal(str(line.unit_price or 0))
            amount += unit_price
        return amount

    def _get_asset_cadastre(self, asset):
        if not asset:
            return ''
        return (getattr(asset, 'land_register', None)
            or getattr(asset, 'home_assessment', None)
            or '')

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

    def _write_back_contract(self, contract):
        contract.cadastre = self.start.cadastre
        contract.document_attribute_set = self.start.attribute_set
        contract.document_attributes = self.start.attributes or {}
        contract.save()

    def _get_render_context(self, contract):
        asset = self.start.asset or self._get_default_asset(contract)
        lessor_company = self.start.lessor_company
        lessor_contact = self.start.lessor_contact
        lessee_company = self.start.lessee_company
        lessee_contacts = list(self.start.lessee_contacts or [])
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
        attachment_names = [
            attachment.name for attachment in self.start.origin_attachments
            if attachment and attachment.name]
        context = SafeDict({
                'today': datetime.now().strftime('%d/%m/%Y'),
                'contract_number': safe_text(contract.number),
                'contract_reference': safe_text(contract.reference),
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
                'lessee_company': TemplateRecord(lessee_company)
                if lessee_company else '',
                'lessee_company_name': safe_text(lessee_company.rec_name
                    if lessee_company else ''),
                'lessee_contacts': wrapped_lessee_contacts,
                'lessee_contacts_text': ', '.join(lessee_contact_names),
                'asset': TemplateRecord(asset) if asset else '',
                'asset_name': safe_text(asset.rec_name if asset else ''),
                'asset_address': self._get_asset_address(asset) or '; '.join(addresses),
                'cadastre': safe_text(self.start.cadastre),
                'amount': safe_text(self.start.amount),
                'start_date': safe_text(contract.start_date),
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
        source = (text or '').replace('\r\n', '\n')
        environment = Environment(
            autoescape=False,
            undefined=ChainableUndefined,
            trim_blocks=False,
            lstrip_blocks=False)
        template = environment.from_string(source)
        return template.render(**context)

    def _build_docx(self, context):
        paragraphs = []
        paragraphs.append({
                'text': 'CONTRACT %s' % (context['contract_number']
                    or context['contract_reference']
                    or '').strip(),
                'bold': True,
                'center': True,
                })
        paragraphs.append({'text': ''})

        self._append_line_section(paragraphs, self.start.parties_title,
            self.start.parties, context)
        self._append_line_section(paragraphs, self.start.appearances_title,
            self.start.appearances, context)

        if self.start.statements_title:
            paragraphs.append({
                    'text': self.start.statements_title,
                    'bold': True,
                    'center': True,
                    })
            for index, statement in enumerate(sorted(self.start.statements,
                    key=lambda l: ((l.sequence is None), l.sequence or 0,
                        l.id or 0)), start=1):
                content = self._render_text(statement.content, context)
                if statement.title:
                    paragraphs.append({
                            'text': '%s. %s' % (index, statement.title),
                            'bold': True,
                            })
                self._append_markdown(paragraphs, content)
                paragraphs.append({'text': ''})

        seen = set()
        ordered_clauses = []
        for line in sorted(self.start.clauses,
                key=lambda l: ((l.sequence is None), l.sequence or 0, l.id or 0)):
            if line.clause:
                self._append_clause_tree(line.clause, ordered_clauses, seen)

        if ordered_clauses and self.start.clauses_title:
            paragraphs.append({'text': ''})
            paragraphs.append({
                    'text': self.start.clauses_title,
                    'bold': True,
                    'center': True,
                    })
            for index, clause in enumerate(ordered_clauses, start=1):
                title = clause.title or clause.name
                paragraphs.append({
                        'text': '%s. %s' % (index, title),
                        'bold': True,
                        })
                rendered = self._render_text(clause.content, context)
                self._append_markdown(paragraphs, rendered)
                paragraphs.append({'text': ''})

        return self._create_docx(paragraphs)

    def _append_line_section(self, paragraphs, title, lines, context):
        if not title:
            return
        paragraphs.append({
                'text': title,
                'bold': True,
                'center': True,
                })
        for section_line in sorted(lines,
                key=lambda l: ((l.sequence is None), l.sequence or 0,
                    l.id or 0)):
            if section_line.title:
                paragraphs.append({
                        'text': section_line.title,
                        'bold': True,
                        })
            rendered = self._render_text(section_line.content, context)
            self._append_markdown(paragraphs, rendered)
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
