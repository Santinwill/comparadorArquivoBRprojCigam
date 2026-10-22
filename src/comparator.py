from collections import defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from typing import Any
from .normalizer import (
    normalize_business_unit,
    normalize_date,
    normalize_document,
    normalize_fiscal_classification,
    normalize_invoice_number,
)
from .matcher import (
    BASE_FISCAL_PATH,
    NF_FISCAL_ENTRADA_PATH,
    analyze_all_records,
    load_records,
)
from .comparison_rules import (
    RULE_COMPARE,
    RULE_REVIEW,
    get_comparison_fields,
    get_review_fields,
    get_rule,
)

IDENTIFICATION_FIELDS = {
    "CD_UNI_NEG",
    "NF",
    "CNPJ_CPF",
    "DATA_EMISSAO",
}


NUMERIC_FIELDS = {
    "QTD_ENTRADA",
    "VLR_ENTRADA",
    "VLR_FRETE",
    "VLR_DESCONTO",
    "BASE_ICMS",
    "VLR_ICMS",
    "ICMS_NAO_TRI",
    "ICMS_OUTROS",
    "ICMS_NAO_DEST",
    "PIS_NAO_DEST",
    "COFINS_NAO_DEST",
    "BASE_PIS_COFINS",
    "VLR_PIS",
    "VLR_COFINS",
    "VLR_IPI",
    "IPI_NAO_TRI",
    "IPI_OUTROS",
    "VLR_PIS_RET",
    "VLR_COFINS_RET",
    "VLR_CSLL_RET",
    "VLR_IRRF_RET",
    "VLR_ISS_RET",
    "VLR_INSS_RET",
    "VLR_DIF_ALIQUOTA",
}


LOW_CONFIDENCE_MATCH_STATUSES = {
    "MATCH_POR_ELIMINACAO",
}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""

    return str(value).strip()


def normalize_numeric_value(value: Any) -> Decimal | str:
    if value is None or value == "":
        return Decimal("0")

    if isinstance(value, bool):
        return str(value)

    try:
        return Decimal(str(value).strip())

    except (InvalidOperation, ValueError):
        return normalize_text(value)


def normalize_field_value(
    field_name: str,
    value: Any,
) -> Any:
    if field_name == "CD_UNI_NEG":
        return normalize_business_unit(value)

    if field_name == "NF":
        return normalize_invoice_number(value)

    if field_name == "CNPJ_CPF":
        return normalize_document(value)

    if field_name in {
        "DATA_EMISSAO",
        "DATA_ENTRADA",
    }:
        return normalize_date(value)

    if field_name == "CD_CLASS_FISCAL":
        return normalize_fiscal_classification(value)

    if field_name in NUMERIC_FIELDS:
        return normalize_numeric_value(value)

    return normalize_text(value)


def values_are_equal(
    field_name: str,
    base_value: Any,
    cigam_value: Any,
) -> bool:

    normalized_base = normalize_field_value(
        field_name,
        base_value,
    )

    normalized_cigam = normalize_field_value(
        field_name,
        cigam_value,
    )

    return normalized_base == normalized_cigam


def format_value_for_display(value: Any) -> Any:
    if isinstance(value, (datetime, date)):
        return value

    return value


def compare_pair(
    match: dict,
) -> dict:
    base_record = match["base"]
    cigam_record = match["cigam"]
    match_status = match["status"]

    differences = []
    review_differences = []

    comparison_fields = get_comparison_fields()
    review_fields = get_review_fields()

    fields_to_analyze = (
        comparison_fields
        + review_fields
    )

    for field_name in fields_to_analyze:

        base_value = base_record.get(field_name)
        cigam_value = cigam_record.get(field_name)

        equal = values_are_equal(
            field_name,
            base_value,
            cigam_value,
        )

        if equal:
            continue

        field_rule = get_rule(field_name)

        difference = {
            "field": field_name,
            "rule": field_rule["rule"],
            "strategy": field_rule["strategy"],
            "description": field_rule["description"],
            "base_value": format_value_for_display(
                base_value
            ),
            "cigam_value": format_value_for_display(
                cigam_value
            ),
            "normalized_base": normalize_field_value(
                field_name,
                base_value,
            ),
            "normalized_cigam": normalize_field_value(
                field_name,
                cigam_value,
            ),
        }

        if field_rule["rule"] == RULE_COMPARE:
            differences.append(difference)

        elif field_rule["rule"] == RULE_REVIEW:
            review_differences.append(difference)

    if differences:
        comparison_status = "DIVERGENTE"

    elif review_differences:
        comparison_status = "REVISAO_PENDENTE"

    else:
        comparison_status = "SEM_DIVERGENCIA"

    if match_status in LOW_CONFIDENCE_MATCH_STATUSES:
        confidence = "BAIXA"

    elif match_status in {
        "MATCH_DUPLICADO_SEQUENCIAL",
        "MATCH_DUPLICADO_SEM_NCM",
        "MATCH_DESEMPATADO_POR_VALOR",
    }:
        confidence = "MEDIA"

    else:
        confidence = "ALTA"

    return {
        "comparison_status": comparison_status,
        "match_status": match_status,
        "match_confidence": confidence,
        "invoice_key": match["invoice_key"],
        "base_row": base_record.get("_EXCEL_ROW"),
        "cigam_row": cigam_record.get("_EXCEL_ROW"),
        "business_unit": cigam_record.get("CD_UNI_NEG"),
        "invoice_number": cigam_record.get("NF"),
        "document": cigam_record.get("CNPJ_CPF"),
        "emission_date": cigam_record.get("DATA_EMISSAO"),
        "base_item_key": base_record.get("CHAVE_ITEM"),
        "base_record": base_record,
        "cigam_record": cigam_record,
        "differences": differences,
        "review_differences": review_differences,
        "different_fields": [
            difference["field"]
            for difference in differences
        ],
        "review_fields": [
            difference["field"]
            for difference in review_differences
        ],
        "difference_count": len(differences),
        "review_difference_count": len(
            review_differences
        ),
    }


def compare_matched_records(
    matching_results: dict,
) -> dict:
    comparisons = []
    divergent = []
    review_pending = []
    without_difference = []

    difference_count_by_field = defaultdict(int)
    review_count_by_field = defaultdict(int)

    for match in matching_results["matches"]:

        comparison = compare_pair(match)

        comparisons.append(comparison)

        if comparison["differences"]:

            divergent.append(comparison)

            for difference in comparison["differences"]:

                field_name = difference["field"]

                difference_count_by_field[
                    field_name
                ] += 1

        elif comparison["review_differences"]:

            review_pending.append(comparison)

        else:
            without_difference.append(comparison)

        for difference in comparison[
            "review_differences"
        ]:

            field_name = difference["field"]

            review_count_by_field[
                field_name
            ] += 1

    return {
        "comparisons": comparisons,
        "divergent": divergent,
        "review_pending": review_pending,
        "without_difference": without_difference,
        "difference_count_by_field": dict(
            sorted(
                difference_count_by_field.items(),
                key=lambda item: (
                    -item[1],
                    item[0],
                ),
            )
        ),
        "review_count_by_field": dict(
            sorted(
                review_count_by_field.items(),
                key=lambda item: (
                    -item[1],
                    item[0],
                ),
            )
        ),
        "invoice_not_found": matching_results[
            "invoice_not_found"
        ],
        "unresolved": matching_results[
            "unresolved"
        ],
        "unused_base": matching_results[
            "unused_base"
        ],
    }


def print_comparison_example(
    comparison: dict,
):
    print()
    print("-" * 100)

    print(
        f"NF: {comparison['invoice_number']!r} | "
        f"CNPJ: {comparison['document']!r}"
    )

    print(
        f"Linha Base: {comparison['base_row']} | "
        f"Linha CIGAM: {comparison['cigam_row']}"
    )

    print(
        f"Matching: {comparison['match_status']} | "
        f"Confiança: {comparison['match_confidence']}"
    )

    print(
        f"Quantidade de campos divergentes: "
        f"{comparison['difference_count']}"
    )

    for difference in comparison["differences"]:

        print(
            f"  {difference['field']}: "
            f"BASE={difference['base_value']!r} | "
            f"CIGAM={difference['cigam_value']!r}"
        )


def print_summary(
    comparison_results: dict,
):
    comparisons = comparison_results["comparisons"]
    divergent = comparison_results["divergent"]

    print()
    print("=" * 100)
    print("RESULTADO DA COMPARAÇÃO")
    print("=" * 100)

    print(
        f"Pares associados analisados: "
        f"{len(comparisons)}"
    )

    print(
        f"Pares sem divergência:       "
        f"{len(comparison_results['without_difference'])}"
    )

    print(
        f"Pares somente para revisão:  "
        f"{len(comparison_results['review_pending'])}"
    )

    print(
        f"Pares com divergência:       "
        f"{len(divergent)}"
    )

    print(
        f"Matches não resolvidos:      "
        f"{len(comparison_results['unresolved'])}"
    )

    print(
        f"Notas não encontradas:       "
        f"{len(comparison_results['invoice_not_found'])}"
    )

    print(
        f"Itens Base não utilizados:   "
        f"{len(comparison_results['unused_base'])}"
    )

    print()
    print("-" * 100)
    print("DIVERGÊNCIAS POR CAMPO")
    print("-" * 100)

    difference_counts = comparison_results[
        "difference_count_by_field"
    ]

    if not difference_counts:
        print("Nenhuma divergência encontrada.")

    else:
        for field_name, count in difference_counts.items():

            print(
                f"{field_name:<30} "
                f"{count}"
            )

    print()
    print("-" * 100)
    print("DIFERENÇAS EM REVISÃO POR CAMPO")
    print("-" * 100)

    review_counts = comparison_results[
        "review_count_by_field"
    ]

    if not review_counts:
        print("Nenhuma diferença em revisão.")

    else:
        for field_name, count in review_counts.items():

            print(
                f"{field_name:<30} "
                f"{count}"
            )            

    print()
    print("=" * 100)
    print("EXEMPLOS DE PARES DIVERGENTES")
    print("=" * 100)

    if not divergent:
        print("Nenhum par divergente.")

    else:
        for comparison in divergent[:10]:

            print_comparison_example(
                comparison
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

    matching_results = analyze_all_records(
        base_records,
        cigam_records,
    )

    print("Comparando campos fiscais...")

    comparison_results = compare_matched_records(
        matching_results
    )

    print_summary(
        comparison_results
    )


if __name__ == "__main__":
    main()