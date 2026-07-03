from trytond.pool import Pool, PoolMeta
from trytond.transaction import Transaction

from .contract import get_sync_group_key, get_unique_records


class Asset(metaclass=PoolMeta):
    __name__ = 'asset'

    @classmethod
    def create(cls, vlist):
        assets = super().create(vlist)
        if Transaction().context.get('skip_asset_attribute_sync'):
            return assets
        to_sync = []
        for asset, values in zip(assets, vlist):
            if cls._has_attribute_changes(values):
                to_sync.append(asset)
        if to_sync:
            cls._sync_attributes_to_contracts(to_sync)
        return assets

    @classmethod
    def write(cls, *args):
        super().write(*args)
        if Transaction().context.get('skip_asset_attribute_sync'):
            return
        to_sync = []
        actions = iter(args)
        for assets, values in zip(actions, actions):
            if cls._has_attribute_changes(values):
                to_sync.extend(assets)
        if to_sync:
            cls._sync_attributes_to_contracts(
                cls.browse(list({a.id for a in to_sync})))

    @classmethod
    def _has_attribute_changes(cls, values):
        return any(name in values for name in ('attribute_set', 'attributes'))

    def get_related_contracts(self):
        return get_unique_records(line.contract for line in self.contract_lines
            if getattr(line, 'contract', None))

    @classmethod
    def _get_contract_sync_values(cls, contract, preferred_asset_ids=None):
        if preferred_asset_ids is None:
            preferred_asset_ids = set()
        assets = contract.get_related_assets()
        preferred_assets = [asset for asset in assets
            if asset.id in preferred_asset_ids]
        other_assets = [asset for asset in assets
            if asset.id not in preferred_asset_ids]
        ordered_assets = other_assets + preferred_assets
        attribute_set = None
        attributes = {}
        for asset in ordered_assets:
            if not attribute_set and getattr(asset, 'attribute_set', None):
                attribute_set = asset.attribute_set
            attributes.update(dict(getattr(asset, 'attributes', None) or {}))
        return {
            'document_attribute_set': attribute_set.id if attribute_set else None,
            'document_attributes': attributes,
            }

    @classmethod
    def _sync_attributes_to_contracts(cls, assets):
        Contract = Pool().get('contract')
        contract_preferred_assets = {}
        for asset in assets:
            contracts = asset.get_related_contracts()
            if not contracts:
                continue
            for contract in contracts:
                contract_preferred_assets.setdefault(contract.id, set()).add(
                    asset.id)
        if not contract_preferred_assets:
            return
        grouped_contracts = {}
        for contract in Contract.browse(list(contract_preferred_assets.keys())):
            values = cls._get_contract_sync_values(contract,
                preferred_asset_ids=contract_preferred_assets[contract.id])
            key = get_sync_group_key(values['document_attribute_set'],
                values['document_attributes'])
            grouped_contracts.setdefault(key, {
                    'values': values,
                    'contracts': [],
                    })
            grouped_contracts[key]['contracts'].append(contract)
        with Transaction().set_context(skip_contract_document_asset_sync=True):
            for data in grouped_contracts.values():
                Contract.write(data['contracts'], data['values'])
