from collections import defaultdict
from pathlib import Path
from typing import Any
from openpyxl import load_workbook
from .normalizer import (
    build_invoice_key,
    build_item_key,
    normalize_date,
)


PROJECT_ROOT = Path(__file__).resolve().parent.parent
INPUT_DIR = PROJECT_ROOT / "input"

BASE_FISCAL_PATH = INPUT_DIR / "Base_Fiscal.xlsx"
NF_FISCAL_ENTRADA_PATH = INPUT_DIR / "NF_FISCAL_ENTRADA.xlsx"

SHEET_NAME = "Export"

EXAMPLE_LIMIT = 10


def load_records(file_path: Path) -> list[dict]:
    if not file_path.exists():
        raise FileNotFoundError(
            f"Arquivo não encontrado: {file_path}"
        )

    workbook = load_workbook(
        filename=file_path,
        read_only=True,
        data_only=True,
    )

    try:
        if SHEET_NAME not in workbook.sheetnames:
            raise ValueError(
                f"Aba '{SHEET_NAME}' não encontrada em "
                f"{file_path.name}"
            )

        worksheet = workbook[SHEET_NAME]

        rows = worksheet.iter_rows(values_only=True)
        headers = next(rows)

        records = []

        for excel_row_number, row in enumerate(rows, start=2):

            record = {}

            for index, header in enumerate(headers):

                if header is None:
                    continue

                # A Base Fiscal possui um cabeçalho duplicado.
                if header in record:
                    continue

                record[header] = row[index]

            record["_EXCEL_ROW"] = excel_row_number
            records.append(record)

        return records

    finally:
        workbook.close()


def build_invoice_date_key(
    record: dict,
) -> tuple[str, str, str, str]:
    invoice_key = build_invoice_key(
        record.get("CD_UNI_NEG"),
        record.get("NF"),
        record.get("CNPJ_CPF"),
    )

    emission_date = normalize_date(
        record.get("DATA_EMISSAO")
    )

    return (
        invoice_key[0],
        invoice_key[1],
        invoice_key[2],
        emission_date,
    )


def create_invoice_index(
    records: list[dict],
) -> dict:
    index = defaultdict(list)

    for record in records:

        invoice_key = build_invoice_date_key(
            record
        )

        index[invoice_key].append(record)

    return index


def normalize_value(value: Any) -> str:
    if value is None:
        return ""

    return str(value).strip()


def build_quantity_value_key(
    record: dict,
) -> tuple[str, str]:
    item_key = build_item_key(
        "",
        record.get("QTD_ENTRADA"),
    )

    quantity = item_key[1]

    value = normalize_value(
        record.get("VLR_ENTRADA")
    )

    return quantity, value


def build_full_item_key(
    record: dict,
) -> tuple[str, str, str]:
    item_key = build_item_key(
        record.get("CD_CLASS_FISCAL"),
        record.get("QTD_ENTRADA"),
    )

    value = normalize_value(
        record.get("VLR_ENTRADA")
    )

    return (
        item_key[0],
        item_key[1],
        value,
    )


def group_indexes_by_key(
    records: list[dict],
    available_indexes: set[int],
    key_function,
) -> dict:
    groups = defaultdict(list)

    for index in sorted(available_indexes):

        key = key_function(records[index])

        groups[key].append(index)

    return groups


def add_match(
    matches: list[dict],
    status: str,
    base_record: dict,
    cigam_record: dict,
    invoice_key: tuple,
):
    matches.append(
        {
            "status": status,
            "invoice_key": invoice_key,
            "base": base_record,
            "cigam": cigam_record,
        }
    )


def match_invoice_items(
    invoice_key: tuple,
    base_items: list[dict],
    cigam_items: list[dict],
) -> dict:
    available_base = set(range(len(base_items)))
    available_cigam = set(range(len(cigam_items)))

    matches = []

    base_groups = group_indexes_by_key(
        base_items,
        available_base,
        lambda record: build_item_key(
            record.get("CD_CLASS_FISCAL"),
            record.get("QTD_ENTRADA"),
        ),
    )

    cigam_groups = group_indexes_by_key(
        cigam_items,
        available_cigam,
        lambda record: build_item_key(
            record.get("CD_CLASS_FISCAL"),
            record.get("QTD_ENTRADA"),
        ),
    )

    common_keys = (
        set(base_groups)
        & set(cigam_groups)
    )

    for item_key in sorted(common_keys):

        base_indexes = base_groups[item_key]
        cigam_indexes = cigam_groups[item_key]

        if (
            len(base_indexes) == 1
            and len(cigam_indexes) == 1
        ):
            base_index = base_indexes[0]
            cigam_index = cigam_indexes[0]

            add_match(
                matches,
                "MATCH_EXATO",
                base_items[base_index],
                cigam_items[cigam_index],
                invoice_key,
            )

            available_base.remove(base_index)
            available_cigam.remove(cigam_index)

    base_groups = group_indexes_by_key(
        base_items,
        available_base,
        build_full_item_key,
    )

    cigam_groups = group_indexes_by_key(
        cigam_items,
        available_cigam,
        build_full_item_key,
    )

    common_keys = (
        set(base_groups)
        & set(cigam_groups)
    )

    for full_key in sorted(common_keys):

        base_indexes = base_groups[full_key]
        cigam_indexes = cigam_groups[full_key]

        pair_count = min(
            len(base_indexes),
            len(cigam_indexes),
        )

        if pair_count == 0:
            continue

        if (
            len(base_indexes) == 1
            and len(cigam_indexes) == 1
        ):
            status = "MATCH_DESEMPATADO_POR_VALOR"
        else:
            status = "MATCH_DUPLICADO_SEQUENCIAL"

        for position in range(pair_count):

            base_index = base_indexes[position]
            cigam_index = cigam_indexes[position]

            if base_index not in available_base:
                continue

            if cigam_index not in available_cigam:
                continue

            add_match(
                matches,
                status,
                base_items[base_index],
                cigam_items[cigam_index],
                invoice_key,
            )

            available_base.remove(base_index)
            available_cigam.remove(cigam_index)

    base_groups = group_indexes_by_key(
        base_items,
        available_base,
        build_quantity_value_key,
    )

    cigam_groups = group_indexes_by_key(
        cigam_items,
        available_cigam,
        build_quantity_value_key,
    )

    common_keys = (
        set(base_groups)
        & set(cigam_groups)
    )

    for quantity_value_key in sorted(common_keys):

        base_indexes = base_groups[
            quantity_value_key
        ]

        cigam_indexes = cigam_groups[
            quantity_value_key
        ]

        pair_count = min(
            len(base_indexes),
            len(cigam_indexes),
        )

        if pair_count == 0:
            continue

        if (
            len(base_indexes) == 1
            and len(cigam_indexes) == 1
        ):
            status = "MATCH_SEM_NCM"

        else:
            status = "MATCH_DUPLICADO_SEM_NCM"

        for position in range(pair_count):

            base_index = base_indexes[position]
            cigam_index = cigam_indexes[position]

            if base_index not in available_base:
                continue

            if cigam_index not in available_cigam:
                continue

            add_match(
                matches,
                status,
                base_items[base_index],
                cigam_items[cigam_index],
                invoice_key,
            )

            available_base.remove(base_index)
            available_cigam.remove(cigam_index)

    if (
        len(available_base) == 1
        and len(available_cigam) == 1
    ):
        base_index = next(iter(available_base))
        cigam_index = next(iter(available_cigam))

        add_match(
            matches,
            "MATCH_POR_ELIMINACAO",
            base_items[base_index],
            cigam_items[cigam_index],
            invoice_key,
        )

        available_base.remove(base_index)
        available_cigam.remove(cigam_index)

    unresolved_cigam = [
        cigam_items[index]
        for index in sorted(available_cigam)
    ]

    unused_base = [
        base_items[index]
        for index in sorted(available_base)
    ]

    return {
        "matches": matches,
        "unresolved_cigam": unresolved_cigam,
        "unused_base": unused_base,
    }


def analyze_all_records(
    base_records: list[dict],
    cigam_records: list[dict],
) -> dict:
    base_index = create_invoice_index(
        base_records
    )

    cigam_index = create_invoice_index(
        cigam_records
    )

    results = {
        "matches": [],
        "invoice_not_found": [],
        "unresolved": [],
        "unused_base": [],
    }

    for invoice_key, cigam_items in cigam_index.items():

        base_items = base_index.get(
            invoice_key,
            [],
        )

        if not base_items:

            for cigam_item in cigam_items:

                results["invoice_not_found"].append(
                    {
                        "status": "NOTA_NAO_ENCONTRADA",
                        "invoice_key": invoice_key,
                        "cigam": cigam_item,
                    }
                )

            continue

        invoice_result = match_invoice_items(
            invoice_key,
            base_items,
            cigam_items,
        )

        results["matches"].extend(
            invoice_result["matches"]
        )

        for cigam_item in invoice_result[
            "unresolved_cigam"
        ]:
            results["unresolved"].append(
                {
                    "status": "MATCH_NAO_RESOLVIDO",
                    "invoice_key": invoice_key,
                    "cigam": cigam_item,
                }
            )

        for base_item in invoice_result[
            "unused_base"
        ]:
            results["unused_base"].append(
                {
                    "status": "ITEM_BASE_NAO_UTILIZADO",
                    "invoice_key": invoice_key,
                    "base": base_item,
                }
            )

    return results


def count_statuses(results: dict) -> dict:
    counts = defaultdict(int)

    for match in results["matches"]:
        counts[match["status"]] += 1

    counts["MATCH_NAO_RESOLVIDO"] = len(
        results["unresolved"]
    )

    counts["NOTA_NAO_ENCONTRADA"] = len(
        results["invoice_not_found"]
    )

    counts["ITEM_BASE_NAO_UTILIZADO"] = len(
        results["unused_base"]
    )

    return counts


def print_match_examples(
    results: dict,
    status: str,
    limit: int = EXAMPLE_LIMIT,
):
    matches = [
        item
        for item in results["matches"]
        if item["status"] == status
    ]

    print()
    print("=" * 100)
    print(f"EXEMPLOS - {status}")
    print("=" * 100)

    if not matches:
        print("Nenhum registro.")
        return

    for item in matches[:limit]:

        base = item["base"]
        cigam = item["cigam"]

        print()

        print(
            f"Linha Base: {base.get('_EXCEL_ROW')} | "
            f"Linha CIGAM: {cigam.get('_EXCEL_ROW')}"
        )

        print(
            f"NF: {cigam.get('NF')!r} | "
            f"CNPJ: {cigam.get('CNPJ_CPF')!r}"
        )

        print(
            f"Base  | "
            f"NCM={base.get('CD_CLASS_FISCAL')!r} | "
            f"QTD={base.get('QTD_ENTRADA')!r} | "
            f"VLR={base.get('VLR_ENTRADA')!r}"
        )

        print(
            f"CIGAM | "
            f"NCM={cigam.get('CD_CLASS_FISCAL')!r} | "
            f"QTD={cigam.get('QTD_ENTRADA')!r} | "
            f"VLR={cigam.get('VLR_ENTRADA')!r}"
        )


def print_unresolved_examples(
    title: str,
    records: list[dict],
    limit: int = EXAMPLE_LIMIT,
):
    print()
    print("=" * 100)
    print(title)
    print("=" * 100)

    if not records:
        print("Nenhum registro.")
        return

    for item in records[:limit]:

        cigam = item["cigam"]

        print()

        print(
            f"Linha CIGAM: "
            f"{cigam.get('_EXCEL_ROW')}"
        )

        print(
            f"NF={cigam.get('NF')!r} | "
            f"CNPJ={cigam.get('CNPJ_CPF')!r} | "
            f"DATA={cigam.get('DATA_EMISSAO')!r}"
        )

        print(
            f"NCM={cigam.get('CD_CLASS_FISCAL')!r} | "
            f"QTD={cigam.get('QTD_ENTRADA')!r} | "
            f"VLR={cigam.get('VLR_ENTRADA')!r}"
        )


def print_unused_base_examples(
    records: list[dict],
    limit: int = EXAMPLE_LIMIT,
):
    print()
    print("=" * 100)
    print("EXEMPLOS - ITENS DA BASE NÃO UTILIZADOS")
    print("=" * 100)

    if not records:
        print("Nenhum registro.")
        return

    for item in records[:limit]:

        base = item["base"]

        print()

        print(
            f"Linha Base: "
            f"{base.get('_EXCEL_ROW')}"
        )

        print(
            f"NF={base.get('NF')!r} | "
            f"CNPJ={base.get('CNPJ_CPF')!r} | "
            f"DATA={base.get('DATA_EMISSAO')!r}"
        )

        print(
            f"CHAVE_ITEM={base.get('CHAVE_ITEM')!r}"
        )

        print(
            f"NCM={base.get('CD_CLASS_FISCAL')!r} | "
            f"QTD={base.get('QTD_ENTRADA')!r} | "
            f"VLR={base.get('VLR_ENTRADA')!r}"
        )


def print_summary(
    base_records: list[dict],
    cigam_records: list[dict],
    results: dict,
):
    counts = count_statuses(results)

    classified_cigam = (
        len(results["matches"])
        + len(results["unresolved"])
        + len(results["invoice_not_found"])
    )

    print()
    print("=" * 100)
    print("RESULTADO DO MATCHING UM-PARA-UM")
    print("=" * 100)

    print(
        f"Registros Base Fiscal: "
        f"{len(base_records)}"
    )

    print(
        f"Registros CIGAM:       "
        f"{len(cigam_records)}"
    )

    print()
    print("-" * 100)

    statuses = [
        "MATCH_EXATO",
        "MATCH_DESEMPATADO_POR_VALOR",
        "MATCH_DUPLICADO_SEQUENCIAL",
        "MATCH_SEM_NCM",
        "MATCH_DUPLICADO_SEM_NCM",
        "MATCH_POR_ELIMINACAO",
        "MATCH_NAO_RESOLVIDO",
        "NOTA_NAO_ENCONTRADA",
    ]

    for status in statuses:

        print(
            f"{status:<35} "
            f"{counts[status]}"
        )

    print()
    print("-" * 100)

    print(
        f"TOTAL CIGAM CLASSIFICADO: "
        f"{classified_cigam}"
    )

    print(
        f"ITENS DA BASE NÃO UTILIZADOS: "
        f"{counts['ITEM_BASE_NAO_UTILIZADO']}"
    )

    if cigam_records:

        matched_count = len(
            results["matches"]
        )

        print()
        print(
            f"Percentual com alguma associação: "
            f"{matched_count / len(cigam_records) * 100:.2f}%"
        )

    if classified_cigam != len(cigam_records):

        print()
        print(
            "ATENÇÃO: o total classificado não corresponde "
            "ao total de registros do CIGAM."
        )


def main():

    print("Carregando Base Fiscal...")

    base_records = load_records(
        BASE_FISCAL_PATH
    )

    print("Carregando CIGAM...")

    cigam_records = load_records(
        NF_FISCAL_ENTRADA_PATH
    )

    print("Executando matching um-para-um...")

    results = analyze_all_records(
        base_records,
        cigam_records,
    )

    print_summary(
        base_records,
        cigam_records,
        results,
    )

    print_match_examples(
        results,
        "MATCH_DESEMPATADO_POR_VALOR",
    )

    print_match_examples(
        results,
        "MATCH_DUPLICADO_SEQUENCIAL",
    )

    print_match_examples(
        results,
        "MATCH_SEM_NCM",
    )

    print_match_examples(
        results,
        "MATCH_DUPLICADO_SEM_NCM",
    )

    print_match_examples(
        results,
        "MATCH_POR_ELIMINACAO",
    )

    print_unresolved_examples(
        "EXEMPLOS - MATCH NÃO RESOLVIDO",
        results["unresolved"],
    )

    print_unresolved_examples(
        "EXEMPLOS - NOTA NÃO ENCONTRADA",
        results["invoice_not_found"],
    )

    print_unused_base_examples(
        results["unused_base"],
    )


if __name__ == "__main__":
    main()