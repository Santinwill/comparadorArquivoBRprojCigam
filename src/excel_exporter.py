from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import (
    Alignment,
    Font,
    PatternFill,
)
from openpyxl.utils import get_column_letter
from .comparison_rules import (
    get_comparison_fields,
    get_review_fields,
)
from .comparator import (
    NUMERIC_FIELDS,
    compare_matched_records,
)
from .matcher import (
    analyze_all_records,
    load_daily_cigam_records,
    load_period_cigam_records,
    load_relevant_base_records,
)


PROJECT_ROOT = Path(__file__).resolve().parent.parent
OUTPUT_DIR = PROJECT_ROOT / "output"
OUTPUT_FILE = OUTPUT_DIR / "Fiscal_Comparison.xlsx"

# ============================================================
# para utilizar peiodo predeterminado comentar o daily e utilizar o monthly, depois ajustar as datas
# ============================================================
# REPORT_MODE = "DAILY"
REPORT_MODE = "MONTHLY"
REPORT_START_DATE = date(2026, 10, 1)
REPORT_END_DATE = date(2026, 10, 31)

INTEGER_NUMERIC_FORMAT = "0"
DECIMAL_NUMERIC_FORMAT = "0.###############################"

HEADER_FILL = PatternFill(
    fill_type="solid",
    fgColor="1F4E78",
)

HEADER_FONT = Font(
    color="FFFFFF",
    bold=True,
)

RED_FILL = PatternFill(
    fill_type="solid",
    fgColor="FFC7CE",
)

YELLOW_FILL = PatternFill(
    fill_type="solid",
    fgColor="FFF2CC",
)

GREEN_FILL = PatternFill(
    fill_type="solid",
    fgColor="C6EFCE",
)

ORANGE_FILL = PatternFill(
    fill_type="solid",
    fgColor="FCE4D6",
)

LOW_CONFIDENCE_FILL = PatternFill(
    fill_type="solid",
    fgColor="F4B183",
)

GRAY_FILL = PatternFill(
    fill_type="solid",
    fgColor="D9E1F2",
)


METADATA_COLUMNS = [
    "COMPARISON_STATUS",
    "MATCH_STATUS",
    "MATCH_CONFIDENCE",
    "BASE_ROW",
    "CIGAM_ROW",
    "CD_UNI_NEG",
    "NF",
    "CNPJ_CPF",
    "DATA_EMISSAO",
    "CHAVE_ITEM",
    "DIFFERENCE_COUNT",
    "REVIEW_DIFFERENCE_COUNT",
    "DIFFERENT_FIELDS",
    "REVIEW_FIELDS",
]


KEY_OUTPUT_FIELDS = [
    "CD_UNI_NEG",
    "NF",
    "CNPJ_CPF",
    "DATA_EMISSAO",
    "CD_CLASS_FISCAL",
    "QTD_ENTRADA",
    "VLR_ENTRADA",
]

TEXT_IDENTIFIER_FIELDS = {
    "CD_UNI_NEG",
    "NF",
    "CNPJ_CPF",
    "CHAVE_ITEM",
    "CD_CLASS_FISCAL",
    "CTS_ICMS",
    "CTS_PIS",
    "CTS_COFINS",
    "CTS_IPI",
}


TEXT_IDENTIFIER_PREFIXES = (
    "BASE_",
    "CIGAM_",
)


def excel_safe_value(value):
    if isinstance(value, Decimal):
        return str(value)

    if isinstance(value, (datetime, date)):
        return value

    return value


def style_header(worksheet):
    for cell in worksheet[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

    worksheet.freeze_panes = "A2"
    worksheet.row_dimensions[1].height = 35

    if worksheet.max_column > 0:
        worksheet.auto_filter.ref = worksheet.dimensions


def adjust_column_widths(
    worksheet,
    maximum_width=35,
):

    for column_cells in worksheet.columns:

        column_index = column_cells[0].column
        column_letter = get_column_letter(column_index)

        maximum_length = 0

        for cell in column_cells:

            value = cell.value

            if value is None:
                continue

            value_length = len(str(value))

            if value_length > maximum_length:
                maximum_length = value_length

        adjusted_width = min(
            maximum_length + 2,
            maximum_width,
        )

        worksheet.column_dimensions[
            column_letter
        ].width = max(adjusted_width, 12)

def format_date_columns(worksheet):
    for column_number in range(
        1,
        worksheet.max_column + 1,
    ):
        header = worksheet.cell(
            row=1,
            column=column_number,
        ).value

        if header != "DATA_EMISSAO":
            continue

        for row_number in range(
            2,
            worksheet.max_row + 1,
        ):
            date_cell = worksheet.cell(
                row=row_number,
                column=column_number,
            )

            if date_cell.value is not None:
                date_cell.number_format = "dd/mm/yyyy"


def get_original_field_name(header):
    if header is None:
        return ""

    header = str(header).strip()

    for prefix in TEXT_IDENTIFIER_PREFIXES:

        if header.startswith(prefix):
            return header[len(prefix):]

    return header


def convert_identifier_to_text(
    field_name,
    value,
):
    if value is None:
        return None

    text_value = str(value).strip()

    if field_name == "CNPJ_CPF":

        digits = "".join(
            character
            for character in text_value
            if character.isdigit()
        )

        if len(digits) == 14:
            return digits.zfill(14)

        if len(digits) == 11:
            return digits.zfill(11)

        return digits

    if field_name == "CD_UNI_NEG":

        if text_value.endswith(".0"):

            possible_integer = text_value[:-2]

            if possible_integer.isdigit():
                return possible_integer

        return text_value

    if field_name == "NF":

        if text_value.endswith(".0"):

            possible_integer = text_value[:-2]

            if possible_integer.isdigit():
                return possible_integer

        return text_value

    if field_name == "CD_CLASS_FISCAL":

        if text_value.endswith(".0"):

            possible_integer = text_value[:-2]

            if possible_integer.isdigit():
                return possible_integer

        return text_value

    return text_value


def format_identifier_columns(worksheet):
    for column_number in range(
        1,
        worksheet.max_column + 1,
    ):
        header = worksheet.cell(
            row=1,
            column=column_number,
        ).value

        original_field_name = get_original_field_name(
            header
        )

        if original_field_name not in TEXT_IDENTIFIER_FIELDS:
            continue

        for row_number in range(
            2,
            worksheet.max_row + 1,
        ):
            cell = worksheet.cell(
                row=row_number,
                column=column_number,
            )

            if cell.value is None:
                continue

            cell.value = convert_identifier_to_text(
                original_field_name,
                cell.value,
            )

            cell.number_format = "@"


def format_numeric_columns(worksheet):
    for column_number in range(
        1,
        worksheet.max_column + 1,
    ):
        header = worksheet.cell(
            row=1,
            column=column_number,
        ).value

        original_field_name = get_original_field_name(
            header
        )

        if original_field_name not in NUMERIC_FIELDS:
            continue

        for row_number in range(
            2,
            worksheet.max_row + 1,
        ):
            cell = worksheet.cell(
                row=row_number,
                column=column_number,
            )

            value = cell.value

            if value is None:
                continue

            if isinstance(value, bool):
                continue

            if not isinstance(
                value,
                (
                    int,
                    float,
                    Decimal,
                ),
            ):
                continue

            if isinstance(value, int):
                cell.number_format = INTEGER_NUMERIC_FORMAT
                continue

            if isinstance(value, Decimal):

                if value == value.to_integral_value():
                    cell.number_format = INTEGER_NUMERIC_FORMAT
                else:
                    cell.number_format = DECIMAL_NUMERIC_FORMAT

                continue

            if isinstance(value, float):

                if value.is_integer():
                    cell.number_format = INTEGER_NUMERIC_FORMAT
                else:
                    cell.number_format = DECIMAL_NUMERIC_FORMAT


def highlight_low_confidence_row(
    worksheet,
    row_number,
):
    red_color = RED_FILL.fgColor.rgb
    yellow_color = YELLOW_FILL.fgColor.rgb

    for column_number in range(
        1,
        worksheet.max_column + 1,
    ):
        cell = worksheet.cell(
            row=row_number,
            column=column_number,
        )

        current_color = cell.fill.fgColor.rgb

        if (
            cell.fill.fill_type == "solid"
            and current_color in {
                red_color,
                yellow_color,
            }
        ):
            continue

        cell.fill = LOW_CONFIDENCE_FILL

def get_all_audit_fields():
    fields = (
        get_comparison_fields()
        + get_review_fields()
    )

    return list(dict.fromkeys(fields))


def build_comparison_headers():
    headers = list(METADATA_COLUMNS)

    for field_name in get_all_audit_fields():

        headers.append(
            f"BASE_{field_name}"
        )

        headers.append(
            f"CIGAM_{field_name}"
        )

    return headers


def write_comparison_header(worksheet):
    headers = build_comparison_headers()

    for column_number, header in enumerate(
        headers,
        start=1,
    ):
        worksheet.cell(
            row=1,
            column=column_number,
            value=header,
        )

    style_header(worksheet)

    return {
        header: index
        for index, header in enumerate(
            headers,
            start=1,
        )
    }


def write_comparison_row(
    worksheet,
    comparison,
    column_map,
):
    row_number = worksheet.max_row + 1

    different_fields = set(
        comparison["different_fields"]
    )

    review_fields = set(
        comparison["review_fields"]
    )

    metadata = {
        "COMPARISON_STATUS": comparison[
            "comparison_status"
        ],
        "MATCH_STATUS": comparison[
            "match_status"
        ],
        "MATCH_CONFIDENCE": comparison[
            "match_confidence"
        ],
        "BASE_ROW": comparison["base_row"],
        "CIGAM_ROW": comparison["cigam_row"],
        "CD_UNI_NEG": comparison[
            "business_unit"
        ],
        "NF": comparison[
            "invoice_number"
        ],
        "CNPJ_CPF": comparison["document"],
        "DATA_EMISSAO": comparison[
            "emission_date"
        ],
        "CHAVE_ITEM": comparison[
            "base_item_key"
        ],
        "DIFFERENCE_COUNT": comparison[
            "difference_count"
        ],
        "REVIEW_DIFFERENCE_COUNT": comparison[
            "review_difference_count"
        ],
        "DIFFERENT_FIELDS": ", ".join(
            comparison["different_fields"]
        ),
        "REVIEW_FIELDS": ", ".join(
            comparison["review_fields"]
        ),
    }

    for header, value in metadata.items():

        worksheet.cell(
            row=row_number,
            column=column_map[header],
            value=excel_safe_value(value),
        )

    base_record = comparison["base_record"]
    cigam_record = comparison["cigam_record"]

    for field_name in get_all_audit_fields():

        base_column = column_map[
            f"BASE_{field_name}"
        ]

        cigam_column = column_map[
            f"CIGAM_{field_name}"
        ]

        base_cell = worksheet.cell(
            row=row_number,
            column=base_column,
            value=excel_safe_value(
                base_record.get(field_name)
            ),
        )

        cigam_cell = worksheet.cell(
            row=row_number,
            column=cigam_column,
            value=excel_safe_value(
                cigam_record.get(field_name)
            ),
        )

        if field_name in different_fields:

            base_cell.fill = RED_FILL
            cigam_cell.fill = RED_FILL

        elif field_name in review_fields:

            base_cell.fill = YELLOW_FILL
            cigam_cell.fill = YELLOW_FILL

    status_cell = worksheet.cell(
        row=row_number,
        column=column_map[
            "COMPARISON_STATUS"
        ],
    )

    if comparison["comparison_status"] == "DIVERGENTE":
        status_cell.fill = RED_FILL

    elif comparison["comparison_status"] == "REVISAO_PENDENTE":
        status_cell.fill = YELLOW_FILL

    else:
        status_cell.fill = GREEN_FILL

    match_status_cell = worksheet.cell(
        row=row_number,
        column=column_map["MATCH_STATUS"],
    )

    if comparison["match_status"] == "MATCH_POR_ELIMINACAO":

        highlight_low_confidence_row(
            worksheet,
            row_number,
        )

        match_status_cell.fill = LOW_CONFIDENCE_FILL
        match_status_cell.font = Font(
            bold=True,
            color="9C0006",
        )


def create_comparison_sheet(
    workbook,
    sheet_name,
    comparisons,
):
    worksheet = workbook.create_sheet(
        title=sheet_name
    )

    column_map = write_comparison_header(
        worksheet
    )

    for comparison in comparisons:

        write_comparison_row(
            worksheet,
            comparison,
            column_map,
        )

    format_date_columns(
        worksheet
    )

    format_identifier_columns(
        worksheet
    )

    format_numeric_columns(
        worksheet
    )

    adjust_column_widths(
        worksheet
    )

    return worksheet


def get_status_reason(
    status,
    item,
):
    if status == "MATCH_NAO_RESOLVIDO":

        return (
            "A nota foi localizada na Base Fiscal, mas o item "
            "não apresentou uma associação segura pelas regras "
            "de matching disponíveis."
        )

    if status == "NOTA_NAO_ENCONTRADA":

        return (
            "Não foi localizada na Base Fiscal uma nota com a "
            "mesma combinação de CD_UNI_NEG, NF, CNPJ_CPF e "
            "DATA_EMISSAO."
        )

    if status == "ITEM_BASE_NAO_UTILIZADO":

        return (
            "O item existe na Base Fiscal, mas não foi utilizado "
            "em nenhuma associação com os registros do CIGAM."
        )

    return (
        "Status sem motivo específico configurado."
    )


def write_simple_records_sheet(
    workbook,
    sheet_name,
    records,
    source_name,
):
    worksheet = workbook.create_sheet(
        title=sheet_name
    )

    headers = [
        "STATUS",
        "MOTIVO_STATUS",
        "SOURCE",
        "EXCEL_ROW",
        *KEY_OUTPUT_FIELDS,
    ]

    for column_number, header in enumerate(
        headers,
        start=1,
    ):
        worksheet.cell(
            row=1,
            column=column_number,
            value=header,
        )

    style_header(
        worksheet
    )

    for item in records:

        record = item[source_name]
        status = item.get("status")

        status_reason = get_status_reason(
            status,
            item,
        )

        row = [
            status,
            status_reason,
            source_name.upper(),
            record.get("_EXCEL_ROW"),
        ]

        for field_name in KEY_OUTPUT_FIELDS:

            row.append(
                excel_safe_value(
                    record.get(field_name)
                )
            )

        worksheet.append(row)

        result_row = worksheet.max_row

        status_cell = worksheet.cell(
            row=result_row,
            column=1,
        )

        reason_cell = worksheet.cell(
            row=result_row,
            column=2,
        )

        status_cell.fill = ORANGE_FILL
        reason_cell.fill = ORANGE_FILL

        reason_cell.alignment = Alignment(
            vertical="top",
            wrap_text=True,
        )

    format_date_columns(
        worksheet
    )

    format_identifier_columns(
        worksheet
    )

    format_numeric_columns(
        worksheet
    )

    adjust_column_widths(
        worksheet
    )

    worksheet.column_dimensions["B"].width = 65

    return worksheet


def add_color_legend(worksheet):
    worksheet["E3"] = "Legenda de cores"
    worksheet["E3"].fill = HEADER_FILL
    worksheet["E3"].font = HEADER_FONT
    worksheet["E3"].alignment = Alignment(
        horizontal="center",
        vertical="center",
    )

    worksheet["F3"] = "Significado"
    worksheet["F3"].fill = HEADER_FILL
    worksheet["F3"].font = HEADER_FONT
    worksheet["F3"].alignment = Alignment(
        horizontal="center",
        vertical="center",
    )

    legend_items = [
        (
            "Vermelho",
            "Divergência confirmada",
            RED_FILL,
        ),
        (
            "Amarelo",
            "Diferença pendente de revisão fiscal",
            YELLOW_FILL,
        ),
        (
            "Laranja",
            "Match por eliminação com confiança baixa",
            LOW_CONFIDENCE_FILL,
        ),
        (
            "Verde",
            "Registro sem divergência",
            GREEN_FILL,
        ),
        (
            "Azul",
            "Cabeçalhos e títulos",
            HEADER_FILL,
        ),
    ]

    start_row = 4

    for position, legend_item in enumerate(
        legend_items,
        start=start_row,
    ):
        color_name, description, fill = legend_item

        color_cell = worksheet.cell(
            row=position,
            column=5,
            value=color_name,
        )

        description_cell = worksheet.cell(
            row=position,
            column=6,
            value=description,
        )

        color_cell.fill = fill
        color_cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

        description_cell.alignment = Alignment(
            vertical="center",
            wrap_text=True,
        )

        if color_name == "Azul":
            color_cell.font = HEADER_FONT

        elif color_name == "Laranja":
            color_cell.font = Font(
                bold=True,
                color="9C0006",
            )

    worksheet.column_dimensions["E"].width = 18
    worksheet.column_dimensions["F"].width = 48

    for row_number in range(
        start_row,
        start_row + len(legend_items),
    ):
        worksheet.row_dimensions[
            row_number
        ].height = 25


def create_summary_sheet(
    workbook,
    base_records,
    cigam_records,
    matching_results,
    comparison_results,
):
    worksheet = workbook.create_sheet(
        title="Resumo",
        index=0,
    )

    worksheet["A1"] = "Resumo da Comparação Fiscal"
    worksheet["A1"].font = Font(
        bold=True,
        size=16,
        color="FFFFFF",
    )
    worksheet["A1"].fill = HEADER_FILL

    worksheet.merge_cells(
        start_row=1,
        start_column=1,
        end_row=1,
        end_column=3,
    )

    worksheet["A3"] = "Métrica"
    worksheet["B3"] = "Quantidade"
    worksheet["C3"] = "Descrição"

    for cell in worksheet[3]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

    worksheet.row_dimensions[3].height = 30

    metrics = [
        (
            "Registros da Base Fiscal",
            len(base_records),
            "Total de registros na Base Fiscal",
        ),
        (
            "Registros do CIGAM",
            len(cigam_records),
            "Total de registros no NF_FISCAL_ENTRADA",
        ),
        (
            "Pares associados",
            len(comparison_results["comparisons"]),
            "Pares associados pelo matching",
        ),
        (
            "Divergências confirmadas",
            len(comparison_results["divergent"]),
            "Pares com divergências confirmadas",
        ),
        (
            "Somente revisão",
            len(comparison_results["review_pending"]),
            "Pares somente com diferenças em revisão",
        ),
        (
            "Sem divergências",
            len(comparison_results["without_difference"]),
            "Pares sem divergências ou pendências",
        ),
        (
            "Matches não resolvidos",
            len(comparison_results["unresolved"]),
            "Itens sem associação segura",
        ),
        (
            "Notas não encontradas",
            len(comparison_results["invoice_not_found"]),
            "Linhas do CIGAM sem nota correspondente",
        ),
        (
            "Itens da Base não utilizados",
            len(comparison_results["unused_base"]),
            "Itens da Base Fiscal não utilizados",
        ),
    ]

    for metric in metrics:
        worksheet.append(metric)

    start_row = worksheet.max_row + 3

    worksheet.cell(
        row=start_row,
        column=1,
        value="Divergências confirmadas por campo",
    )

    worksheet.cell(
        row=start_row,
        column=1,
    ).font = Font(
        bold=True,
        color="FFFFFF",
    )

    worksheet.cell(
        row=start_row,
        column=1,
    ).fill = HEADER_FILL

    row_number = start_row + 1

    for field_name, count in comparison_results[
        "difference_count_by_field"
    ].items():

        worksheet.cell(
            row=row_number,
            column=1,
            value=field_name,
        )

        worksheet.cell(
            row=row_number,
            column=2,
            value=count,
        )

        worksheet.cell(
            row=row_number,
            column=1,
        ).fill = RED_FILL

        row_number += 1

    row_number += 2

    worksheet.cell(
        row=row_number,
        column=1,
        value="Diferenças em revisão por campo",
    )

    worksheet.cell(
        row=row_number,
        column=1,
    ).font = Font(
        bold=True,
        color="FFFFFF",
    )

    worksheet.cell(
        row=row_number,
        column=1,
    ).fill = HEADER_FILL

    row_number += 1

    for field_name, count in comparison_results[
        "review_count_by_field"
    ].items():

        worksheet.cell(
            row=row_number,
            column=1,
            value=field_name,
        )

        worksheet.cell(
            row=row_number,
            column=2,
            value=count,
        )

        worksheet.cell(
            row=row_number,
            column=1,
        ).fill = YELLOW_FILL

        row_number += 1

    worksheet.column_dimensions["A"].width = 35
    worksheet.column_dimensions["B"].width = 15
    worksheet.column_dimensions["C"].width = 50

    add_color_legend(
        worksheet
    )

    worksheet.freeze_panes = "A4"

    return worksheet


def create_workbook(
    base_records,
    cigam_records,
    matching_results,
    comparison_results,
):
    workbook = Workbook()

    default_sheet = workbook.active
    workbook.remove(default_sheet)

    create_summary_sheet(
        workbook,
        base_records,
        cigam_records,
        matching_results,
        comparison_results,
    )

    create_comparison_sheet(
        workbook,
        "Divergencias",
        comparison_results["divergent"],
    )

    create_comparison_sheet(
        workbook,
        "Revisao",
        comparison_results["review_pending"],
    )

    write_simple_records_sheet(
        workbook,
        "Nao_Resolvidos",
        comparison_results["unresolved"],
        "cigam",
    )

    write_simple_records_sheet(
        workbook,
        "Notas_Nao_Encontradas",
        comparison_results["invoice_not_found"],
        "cigam",
    )

    write_simple_records_sheet(
        workbook,
        "Base_Nao_Utilizada",
        comparison_results["unused_base"],
        "base",
    )

    return workbook


def main():
    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    print()
    print("=" * 80)
    print("CONFIGURAÇÃO DO RELATÓRIO")
    print("=" * 80)

    if REPORT_MODE == "DAILY":
        print("Modo de processamento: DIÁRIO")
        print(
            "Carregando registros do CIGAM "
            "referentes ao dia anterior..."
        )

        cigam_records, target_date = (
            load_daily_cigam_records()
        )

        period_description = (
            target_date.strftime("%d/%m/%Y")
        )

        print(
            f"Data de entrada considerada: "
            f"{period_description}"
        )

    elif REPORT_MODE == "MONTHLY":
        print("Modo de processamento: MENSAL")
        print(
            "Carregando registros do CIGAM "
            "referentes ao período informado..."
        )

        (
            cigam_records,
            start_date,
            end_date,
        ) = load_period_cigam_records(
            REPORT_START_DATE,
            REPORT_END_DATE,
        )

        period_description = (
            f"{start_date.strftime('%d/%m/%Y')} "
            f"até "
            f"{end_date.strftime('%d/%m/%Y')}"
        )

        print(
            f"Período de entrada considerado: "
            f"{period_description}"
        )

    else:
        raise ValueError(
            f"REPORT_MODE inválido: "
            f"{REPORT_MODE!r}. "
            f"Utilize 'DAILY' ou 'MONTHLY'."
        )

    print(
        f"Registros CIGAM encontrados: "
        f"{len(cigam_records)}"
    )

    if cigam_records:
        print(
            "Carregando somente os registros relevantes "
            "da Base Fiscal..."
        )

        base_records = (
            load_relevant_base_records(
                cigam_records
            )
        )

        print(
            f"Registros relevantes da Base Fiscal: "
            f"{len(base_records)}"
        )

    else:
        print(
            "Nenhum registro encontrado "
            "para o período."
        )

        print(
            "A Base Fiscal não será carregada."
        )

        base_records = []

    print("Executando matching...")

    matching_results = analyze_all_records(
        base_records,
        cigam_records,
    )

    print("Comparando campos...")

    comparison_results = compare_matched_records(
        matching_results
    )

    print("Criando arquivo Excel...")

    workbook = create_workbook(
        base_records,
        cigam_records,
        matching_results,
        comparison_results,
    )

    workbook.save(
        OUTPUT_FILE
    )

    print()
    print("=" * 80)
    print("ARQUIVO GERADO")
    print("=" * 80)
    print(OUTPUT_FILE)
    print()

    print(
        f"Modo processado: "
        f"{REPORT_MODE}"
    )

    print(
        f"Período processado: "
        f"{period_description}"
    )

    print(
        f"Registros CIGAM processados: "
        f"{len(cigam_records)}"
    )

    print(
        f"Differences: "
        f"{len(comparison_results['divergent'])}"
    )

    print(
        f"Review: "
        f"{len(comparison_results['review_pending'])}"
    )

    print(
        f"Unresolved: "
        f"{len(comparison_results['unresolved'])}"
    )

    print(
        f"Invoice not found: "
        f"{len(comparison_results['invoice_not_found'])}"
    )

    print(
        f"Unused Base: "
        f"{len(comparison_results['unused_base'])}"
    )


if __name__ == "__main__":
    main()