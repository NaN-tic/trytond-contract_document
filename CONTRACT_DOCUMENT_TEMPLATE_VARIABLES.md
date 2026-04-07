# Contract Document Template Variables

This file lists the Jinja2 variables currently available in the `contract_document`
templates.

The same context is available in:
- Parties
- Appearances
- Statements
- Clauses

## Object Variables

These variables are exposed as wrapped records, so you can access fields like
`{{ lessor_contact.rec_name }}`.

- `contract_party`
- `company`
- `lessor_company`
- `lessor_contact`
- `payment_type`
- `bank_account`
- `lessee_company`
- `asset`
- `attribute_set`
- `certificate`

## List Variables

These variables can be iterated.

- `lessee_contacts`
- `attributes`
- `origin_attachments`

Examples:

```jinja2
{% for contact in lessee_contacts %}
- {{ contact.rec_name }}
{% endfor %}
```

```jinja2
{% for key, value in attributes.items() %}
- {{ key }}: {{ value }}
{% endfor %}
```

```jinja2
{% for attachment in origin_attachments %}
- {{ attachment }}
{% endfor %}
```

## Text Variables

These variables are already converted to plain text and can be printed directly.

- `today`
- `contract_number`
- `contract_reference`
- `contract_party_name`
- `company_name`
- `lessor_company_name`
- `lessor_contact_name`
- `payment_type_name`
- `bank_account_name`
- `lessee_company_name`
- `lessee_contacts_text`
- `start_date_text`
- `end_date_text`
- `contract_years_text`
- `asset_name`
- `asset_address`
- `deposit`
- `guarantee_amount`
- `amount`
- `cadastre`
- `home_assessment`
- `energy_certificate`
- `attribute_set_name`
- `certificate_name`
- `first_invoice_date`
- `currency`
- `origin_attachments_text`
- `attributes_block`
- `attachments_block`

## Date, Numeric and Raw Value Variables

These variables keep their original value type where possible.

- `start_date`
- `end_date`
- `contract_years`
- `deposit_value`
- `guarantee_amount_value`
- `amount_value`

## Boolean Variables

- `sign_digitally`

## Dynamic Attribute Variables

Each asset/contract attribute also generates a dynamic variable:

- `attribute_<key>`

Examples:

- `attribute_surface`
- `attribute_rooms`
- `attribute_elevator`

## Common Usage Examples

```jinja2
The contract {{ contract_number }} starts on {{ start_date_text }}
and ends on {{ end_date_text }}.
```

```jinja2
The asset is located at {{ asset_address }}.
```

```jinja2
Lessor: {{ lessor_company.rec_name }}
Contact: {{ lessor_contact.rec_name }}
```

```jinja2
Monthly amount: {{ amount }}
Deposit: {{ deposit }}
Guarantee amount: {{ guarantee_amount }}
```

```jinja2
{% if sign_digitally %}
This document will be digitally signed.
{% endif %}
```

## Notes

- `lessee_contacts` contains record objects and supports field access.
- `origin_attachments` contains attachment names as strings.
- `attributes` is a dictionary.
- If a variable has no value, it is rendered as an empty value by Jinja2.
