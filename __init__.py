# The COPYRIGHT file at the top level of this repository contains
# the full copyright notices and license terms.
from trytond.pool import Pool

from . import contract


def register():
    Pool.register(
        contract.Contract,
        contract.ContractClause,
        contract.ContractParty,
        contract.ContractAppearance,
        contract.ContractManifest,
        contract.ContractBase,
        contract.ContractBaseClause,
        contract.ContractGenerateStart,
        contract.ContractGenerateClause,
        contract.ContractGenerateParty,
        contract.ContractGenerateAppearance,
        contract.ContractGenerateStatement,
        contract.ContractGenerateAttachment,
        module='contract_document', type_='model')
    Pool.register(
        contract.ContractGenerateWizard,
        module='contract_document', type_='wizard')
