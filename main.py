import sys
from pathlib import Path
from src.excel_exporter import main as generate_fiscal_report


PROJECT_ROOT = Path(__file__).resolve().parent

OUTPUT_FILE = (
    PROJECT_ROOT
    / "output"
    / "Fiscal_Comparison.xlsx"
)

FILES_TO_DELETE_AFTER_SUCCESS = [
    PROJECT_ROOT
    / "input"
    / "NF_FISCAL_ENTRADA.xlsx",
]


def delete_previous_report() -> None:
    if not OUTPUT_FILE.exists():
        print(
            "Nenhum relatório anterior encontrado."
        )
        return

    print(
        "Excluindo relatório anterior..."
    )

    OUTPUT_FILE.unlink()

    print(
        f"Relatório anterior excluído: "
        f"{OUTPUT_FILE.name}"
    )


def delete_input_files_after_success() -> None:
    print()
    print("Limpando arquivos de entrada processados...")

    for file_path in FILES_TO_DELETE_AFTER_SUCCESS:

        if not file_path.exists():

            print(
                f"Arquivo não encontrado, nenhuma exclusão: "
                f"{file_path.name}"
            )

            continue

        file_path.unlink()

        print(
            f"Arquivo excluído: "
            f"{file_path.name}"
        )


def main() -> int:
    print()
    print("=" * 80)
    print("COMPARADOR BASE FISCAL X CIGAM")
    print("=" * 80)
    print()

    try:
        delete_previous_report()

        print()

        generate_fiscal_report()

        #
        # Esta função só será executada se todo o processamento
        # anterior terminar sem levantar exceção.
        #
        delete_input_files_after_success()

    except FileNotFoundError as error:

        print()
        print("=" * 80)
        print("ARQUIVO DE ENTRADA NÃO ENCONTRADO")
        print("=" * 80)
        print(error)
        print()
        print(
            "Os arquivos de entrada não foram excluídos."
        )

        return 1

    except PermissionError as error:

        print()
        print("=" * 80)
        print("ARQUIVO BLOQUEADO")
        print("=" * 80)
        print(error)
        print()
        print(
            "Feche os arquivos Excel de entrada e saída "
            "antes de executar novamente."
        )
        print(
            "Os arquivos de entrada não foram excluídos."
        )

        return 1

    except ValueError as error:

        print()
        print("=" * 80)
        print("ERRO DE VALIDAÇÃO")
        print("=" * 80)
        print(error)
        print()
        print(
            "Os arquivos de entrada não foram excluídos."
        )

        return 1

    except Exception as error:

        print()
        print("=" * 80)
        print("ERRO INESPERADO")
        print("=" * 80)

        print(
            f"{type(error).__name__}: {error}"
        )

        print()
        print(
            "Os arquivos de entrada não foram excluídos."
        )

        return 1

    print()
    print("=" * 80)
    print("PROCESSO CONCLUÍDO COM SUCESSO")
    print("=" * 80)

    return 0


if __name__ == "__main__":
    sys.exit(main())