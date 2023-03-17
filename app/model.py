from dataclasses import dataclass
from datetime import datetime
from logging import getLogger
from typing import Any, List, Mapping, NewType, Optional, Sequence
import itertools


logger = getLogger(__name__)


# ProductInput = NewType('ProductInput', Mapping[str, Any])
TableKV = NewType('TableKV', Mapping[str, Any])
JikkinKV = NewType('JikkinKV', Mapping[str, Any])


class ProductInput(Sequence):

    product_kv: Mapping[str, Any]

    def __init__(self, product_kv):

        self.product_kv = product_kv

        self._init_joint_guarantors()

    def _init_joint_guarantors(self):

        joint_guarantors_separator = '*'

        def _split(content: str) -> List[str]:
            """
            連帯保証人用の区切り関数。str.splitだと空文字に対して['']ができてしまうため、
            空文字をsplitしようとしたときに[]になるようにしてある
            """
            return '' if content == '' else content.split(joint_guarantors_separator)

        raw_names = self.product_kv.get("連帯保証人名") or ""

        names = [item.strip() for item in _split(raw_names)]

        fields = None

        if self.product_kv["連帯保証人住所出力"] == "しない":
            fields = [
                names,
                list(itertools.repeat('', len(names))),
                list(itertools.repeat('', len(names))),
            ]
        else:
            fields = [
                names,
                *[[item.strip() for item in _split(field)]
                    for field in [self.product_kv.get("連帯保証人郵便番号") or "",
                                  self.product_kv.get("連帯保証人住所") or "",
                                  ]
                  ]
            ]

        assert fields is not None

        if not all(len(fields[0]) == len(field) for field in fields):
            raise RuntimeError('連帯保証人の入力項目が不正です。修正した上で再度修正してください。')

        self._joint_guarantors = [
            JointGuarantor(*args) for args in zip(*fields)]

    def __getitem__(self, k: str) -> Any:
        return self.product_kv[k]

    def __len__(self) -> int:
        return self.product_kv.__len__()

    def get(self, key: str, *args, **kwargs) -> Any:
        return self.product_kv.get(key, *args, **kwargs)

    def __contains__(self, x: object) -> bool:
        return self.product_kv.__contains__(x)

    @property
    def name(self) -> str:
        return self.product_kv['商品区分']

    @property
    def property_name(self) -> str:
        "物件名"
        return self.product_kv['担保明細－物件名']

    @property
    def state(self) -> str:
        return self.product_kv['担保明細－州国'][0:2]

    @property
    def contract_date(self) -> datetime:
        return self.product_kv['金消契約日']

    @property
    def customer_name(self) -> str:
        """
        顧客名は個人と法人で見るフィールドが違う
        """
        return self.product_kv.get('顧客名') or self.product_kv['法人名']

    @property
    def fiance(self) -> Optional[str]:
        """
        配偶者名
        """
        return self.product_kv.get('配偶者名')

    @property
    def joint_guarantors(self):
        return self._joint_guarantors

    @property
    def is_personal(self):
        return '顧客名' in self.product_kv.keys()


@dataclass(frozen=True)
class JointGuarantor:
    name: str = ''
    postal_code: str = ''
    address: str = ''


@dataclass(frozen=True)
class Product:
    product_input: ProductInput
    jikkin_kv: JikkinKV
    table_kv: TableKV
    jikkin_path: str

    @property
    def name(self) -> str:
        return self.product_input.name

    @property
    def property_name(self) -> str:
        "物件名"
        return self.product_input.property_name

    @property
    def state(self) -> str:
        return self.product_input.state

    @property
    def contract_date(self) -> datetime:
        return self.product_input.contract_date

    @property
    def customer_name(self) -> str:
        """
        顧客名は個人と法人で見るフィールドが違う
        """
        return self.product_input.customer_name

    @property
    def fiance(self) -> Optional[str]:
        """
        配偶者名
        """
        return self.product_input.fiance

    @property
    def joint_guarantors(self):
        return self.product_input.joint_guarantors

    @property
    def is_personal(self) -> bool:
        return self.product_input.is_personal
