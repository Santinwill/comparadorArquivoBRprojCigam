
"""
Ponto principal de execução do comparador fiscal.

Fluxo:
    1. Carrega os arquivos de entrada.
    2. Executa o matching um-para-um.
    3. Compara os campos fiscais.
    4. Gera o relatório Excel.
"""

import sys

from src.excel_exporter import main as generate_fiscal_report


def main() -> int:
    """
    Executa o processo completo de comparação fiscal.

    Retorna:
        0 quando a execução termina corretamente.
        1 quando ocorre algum erro.
    """

    print()
    print("=" * 80)
    print("COMPARADOR BASE FISCAL X CIGAM")
    print("=" * 80)
    print()

    try:
        generate_fiscal_report()

    except FileNotFoundError as error:
        print()
        print("=" * 80)
        print("ARQUIVO DE ENTRADA NÃO ENCONTRADO")
        print("=" * 80)
        print(error)
        print()
        print(
            "Confirme se os arquivos reais estão disponíveis "
            "na pasta input."
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

        return 1

    except ValueError as error:
        print()
        print("=" * 80)
        print("ERRO DE VALIDAÇÃO")
        print("=" * 80)
        print(error)

        return 1

    except Exception as error:
        print()
        print("=" * 80)
        print("ERRO INESPERADO")
        print("=" * 80)
        print(
            f"{type(error).__name__}: {error}"
        )

        return 1

    print()
    print("=" * 80)
    print("PROCESSO CONCLUÍDO COM SUCESSO")
    print("=" * 80)

    return 0


if __name__ == "__main__":
    sys.exit(main())