RULE_COMPARE = "COMPARE"
RULE_REVIEW = "REVIEW"
RULE_IGNORE = "IGNORE"

STRATEGY_TEXT = "TEXT"
STRATEGY_NUMERIC_EXACT = "NUMERIC_EXACT"
STRATEGY_NCM = "NCM"


COMPARISON_RULES = {

    "CD_CLASS_FISCAL": {
        "rule": RULE_COMPARE,
        "strategy": STRATEGY_NCM,
        "description": "NCM / classificação fiscal",
    },

    "QTD_ENTRADA": {
        "rule": RULE_COMPARE,
        "strategy": STRATEGY_NUMERIC_EXACT,
        "description": "Quantidade de entrada",
    },

    "VLR_ENTRADA": {
        "rule": RULE_COMPARE,
        "strategy": STRATEGY_NUMERIC_EXACT,
        "description": "Valor de entrada",
    },

    "VLR_FRETE": {
        "rule": RULE_COMPARE,
        "strategy": STRATEGY_NUMERIC_EXACT,
        "description": "Valor do frete",
    },

    "VLR_DESCONTO": {
        "rule": RULE_COMPARE,
        "strategy": STRATEGY_NUMERIC_EXACT,
        "description": "Valor do desconto",
    },

    "BASE_ICMS": {
        "rule": RULE_COMPARE,
        "strategy": STRATEGY_NUMERIC_EXACT,
        "description": "Base de cálculo do ICMS",
    },

    "VLR_ICMS": {
        "rule": RULE_COMPARE,
        "strategy": STRATEGY_NUMERIC_EXACT,
        "description": "Valor do ICMS",
    },

    "CTS_ICMS": {
        "rule": RULE_REVIEW,
        "strategy": STRATEGY_TEXT,
        "description": "Código tributário do ICMS",
    },

    "ICMS_NAO_TRI": {
        "rule": RULE_REVIEW,
        "strategy": STRATEGY_NUMERIC_EXACT,
        "description": "ICMS não tributado",
    },

    "ICMS_OUTROS": {
        "rule": RULE_REVIEW,
        "strategy": STRATEGY_NUMERIC_EXACT,
        "description": "ICMS outros",
    },

    "ICMS_NAO_DEST": {
        "rule": RULE_REVIEW,
        "strategy": STRATEGY_NUMERIC_EXACT,
        "description": "ICMS não destacado",
    },

    "BASE_PIS_COFINS": {
        "rule": RULE_COMPARE,
        "strategy": STRATEGY_NUMERIC_EXACT,
        "description": "Base de cálculo PIS/COFINS",
    },

    "VLR_PIS": {
        "rule": RULE_COMPARE,
        "strategy": STRATEGY_NUMERIC_EXACT,
        "description": "Valor do PIS",
    },

    "VLR_COFINS": {
        "rule": RULE_COMPARE,
        "strategy": STRATEGY_NUMERIC_EXACT,
        "description": "Valor do COFINS",
    },

    "CTS_PIS": {
        "rule": RULE_REVIEW,
        "strategy": STRATEGY_TEXT,
        "description": "Código tributário do PIS",
    },

    "CTS_COFINS": {
        "rule": RULE_REVIEW,
        "strategy": STRATEGY_TEXT,
        "description": "Código tributário do COFINS",
    },

    "PIS_NAO_DEST": {
        "rule": RULE_REVIEW,
        "strategy": STRATEGY_NUMERIC_EXACT,
        "description": "PIS não destacado",
    },

    "COFINS_NAO_DEST": {
        "rule": RULE_REVIEW,
        "strategy": STRATEGY_NUMERIC_EXACT,
        "description": "COFINS não destacado",
    },

    "VLR_IPI": {
        "rule": RULE_COMPARE,
        "strategy": STRATEGY_NUMERIC_EXACT,
        "description": "Valor do IPI",
    },

    "CTS_IPI": {
        "rule": RULE_REVIEW,
        "strategy": STRATEGY_TEXT,
        "description": "Código tributário do IPI",
    },

    "IPI_NAO_TRI": {
        "rule": RULE_REVIEW,
        "strategy": STRATEGY_NUMERIC_EXACT,
        "description": "IPI não tributado",
    },

    "IPI_OUTROS": {
        "rule": RULE_REVIEW,
        "strategy": STRATEGY_NUMERIC_EXACT,
        "description": "IPI outros",
    },

    "VLR_PIS_RET": {
        "rule": RULE_COMPARE,
        "strategy": STRATEGY_NUMERIC_EXACT,
        "description": "PIS retido",
    },

    "VLR_COFINS_RET": {
        "rule": RULE_COMPARE,
        "strategy": STRATEGY_NUMERIC_EXACT,
        "description": "COFINS retido",
    },

    "VLR_CSLL_RET": {
        "rule": RULE_COMPARE,
        "strategy": STRATEGY_NUMERIC_EXACT,
        "description": "CSLL retida",
    },

    "VLR_IRRF_RET": {
        "rule": RULE_COMPARE,
        "strategy": STRATEGY_NUMERIC_EXACT,
        "description": "IRRF retido",
    },

    "VLR_ISS_RET": {
        "rule": RULE_COMPARE,
        "strategy": STRATEGY_NUMERIC_EXACT,
        "description": "ISS retido",
    },

    "VLR_INSS_RET": {
        "rule": RULE_COMPARE,
        "strategy": STRATEGY_NUMERIC_EXACT,
        "description": "INSS retido",
    },

    "VLR_DIF_ALIQUOTA": {
        "rule": RULE_COMPARE,
        "strategy": STRATEGY_NUMERIC_EXACT,
        "description": "Valor diferencial de alíquota",
    },
}
def get_rule(field_name: str):
    return COMPARISON_RULES.get(field_name)


def get_fields_by_rule(rule_type: str):
    return [
        field_name
        for field_name, config in COMPARISON_RULES.items()
        if config.get("rule") == rule_type
    ]


def get_comparison_fields():
    return get_fields_by_rule(RULE_COMPARE)


def get_review_fields():
    return get_fields_by_rule(RULE_REVIEW)


def get_ignored_fields():
    return get_fields_by_rule(RULE_IGNORE)


def validate_rules():
    valid_rule_types = {
        RULE_COMPARE,
        RULE_REVIEW,
        RULE_IGNORE,
    }

    valid_strategies = {
        STRATEGY_TEXT,
        STRATEGY_NUMERIC_EXACT,
        STRATEGY_NCM,
    }

    for field_name, config in COMPARISON_RULES.items():

        if "rule" not in config:
            raise ValueError(
                f"Campo {field_name!r} não possui a propriedade 'rule'."
            )

        if "strategy" not in config:
            raise ValueError(
                f"Campo {field_name!r} não possui a propriedade "
                f"'strategy'."
            )

        if config["rule"] not in valid_rule_types:
            raise ValueError(
                f"Campo {field_name!r} possui regra inválida: "
                f"{config['rule']!r}"
            )

        if config["strategy"] not in valid_strategies:
            raise ValueError(
                f"Campo {field_name!r} possui estratégia inválida: "
                f"{config['strategy']!r}"
            )


def print_fields(title: str, fields):
    print()
    print(title)
    print("-" * 70)

    if not fields:
        print("  Nenhum")
        return

    for field_name in fields:

        config = get_rule(field_name)

        description = config.get(
            "description",
            "",
        )

        strategy = config.get(
            "strategy",
            "",
        )

        print(
            f"  {field_name:<25} "
            f"{strategy:<15} "
            f"{description}"
        )


def main():
    validate_rules()

    comparison_fields = get_comparison_fields()
    review_fields = get_review_fields()
    ignored_fields = get_ignored_fields()

    print("REGRAS DE COMPARAÇÃO")
    print("=" * 70)

    print_fields(
        "COMPARAÇÃO AUTOMÁTICA:",
        comparison_fields,
    )

    print_fields(
        "REVISÃO PENDENTE:",
        review_fields,
    )

    print_fields(
        "IGNORADOS:",
        ignored_fields,
    )

    print()
    print("=" * 70)

    print(
        f"Total de campos configurados: "
        f"{len(COMPARISON_RULES)}"
    )

    print(
        f"Comparação automática: "
        f"{len(comparison_fields)}"
    )

    print(
        f"Revisão pendente: "
        f"{len(review_fields)}"
    )

    print(
        f"Ignorados: "
        f"{len(ignored_fields)}"
    )


if __name__ == "__main__":
    main()
