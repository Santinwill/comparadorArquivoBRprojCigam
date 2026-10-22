from datetime import date, datetime
import re
from typing import Any


def normalize_business_unit(value: Any) -> str:
    if value is None:
        return ""

    value = str(value).strip()

    if not value:
        return ""

    try:
        return str(int(value))
    except ValueError:
        return value


def normalize_invoice_number(value: Any) -> str:
    if value is None:
        return ""

    value = str(value).strip()

    if not value:
        return ""

    # Trata eventual leitura como 123.0
    if value.endswith(".0"):
        possible_integer = value[:-2]

        if possible_integer.isdigit():
            return possible_integer

    return value


def normalize_document(value: Any) -> str:
    if value is None:
        return ""

    value = str(value).strip()

    if not value:
        return ""

    return re.sub(r"\D", "", value)


def build_invoice_key(
    business_unit: Any,
    invoice_number: Any,
    document: Any,
) -> tuple[str, str, str]:
    return (
        normalize_business_unit(business_unit),
        normalize_invoice_number(invoice_number),
        normalize_document(document),
    )

def normalize_date(value: Any) -> str:
    if value is None:
        return ""

    if isinstance(value, datetime):
        return value.date().isoformat()

    if isinstance(value, date):
        return value.isoformat()

    value = str(value).strip()

    if not value:
        return ""

    return value


def normalize_fiscal_classification(value: Any) -> str:
    if value is None:
        return ""

    digits = re.sub(r"\D", "", str(value))

    if not digits:
        return ""

    if len(digits) >= 8:
        return digits[-8:]

    return digits.zfill(8)


def normalize_quantity(value: Any) -> str:
    if value is None:
        return ""

    return str(value).strip()


def build_item_key(
    fiscal_classification: Any,
    quantity: Any,
) -> tuple[str, str]:
    return (
        normalize_fiscal_classification(
            fiscal_classification
        ),
        normalize_quantity(quantity),
    )

if __name__ == "__main__":

    print("TESTES DO NORMALIZER")
    print("=" * 50)

    print("\nCD_UNI_NEG:")
    print('"016" ->', normalize_business_unit("016"))
    print('"001" ->', normalize_business_unit("001"))
    print("16    ->", normalize_business_unit(16))
    print("1     ->", normalize_business_unit(1))

    print("\nNF:")
    print('"190726" ->', normalize_invoice_number("190726"))
    print("190726   ->", normalize_invoice_number(190726))

    print("\nCNPJ/CPF:")
    print(
        '"06.066.529/0001-94" ->',
        normalize_document("06.066.529/0001-94"),
    )

    print(
        '"06066529000194" ->',
        normalize_document("06066529000194"),
    )

    print("\nCHAVE:")

    base_key = build_invoice_key(
        "016",
        "190726",
        "06066529000194",
    )

    cigam_key = build_invoice_key(
        16,
        190726,
        "06066529000194",
    )

    print("\nDATA:")

    base_date = datetime(
        2026,
        9,
        25,
        17,
        0,
        13,
    )

    cigam_date = datetime(
        2026,
        9,
        25,
        0,
        0,
        0,
    )

    print(
        "Base Fiscal:",
        normalize_date(base_date),
    )

    print(
        "CIGAM:      ",
        normalize_date(cigam_date),
    )

    print(
        "São iguais: ",
        normalize_date(base_date)
        == normalize_date(cigam_date),
    )

    print("\nNCM:")

    print(
        "Base Fiscal:",
        normalize_fiscal_classification(
            "02013000"
        ),
    )

    print(
        "CIGAM:      ",
        normalize_fiscal_classification(
            7002013000
        ),
    )

    print(
        "São iguais: ",
        normalize_fiscal_classification(
            "02013000"
        )
        == normalize_fiscal_classification(
            7002013000
        ),
    )

    print("\nCHAVE DO ITEM:")

    base_item_key = build_item_key(
        "02013000",
        3.028,
    )

    cigam_item_key = build_item_key(
        7002013000,
        3.028,
    )

    print("Base Fiscal:", base_item_key)
    print("CIGAM:      ", cigam_item_key)
    print(
        "São iguais: ",
        base_item_key == cigam_item_key,
    )