"""Testes unitários das regras de negócio (funções puras)."""

from datetime import date

import pytest

import main


# --------------------------------------------------------------------- CPF/PIS

@pytest.mark.parametrize(
    "valor, normalizado",
    [
        ("12345678901", "12345678901"),     # 11 dígitos: mantido
        ("123." + "456." + "789-01", "12345678901"),  # máscara removida
        ("1234567890", None),               # 10 dígitos: rejeitado
        ("123456789012", None),             # 12 dígitos: rejeitado (CPF)
        ("", None),
        (None, None),
        ("abc", None),
    ],
)
def test_cpf_aceita_somente_11_digitos(valor, normalizado):
    assert main.limpar_cpf(valor) == normalizado


@pytest.mark.parametrize(
    "valor, normalizado",
    [
        ("12345678901", "12345678901"),    # 11 dígitos: mantido
        ("912345678901", "12345678901"),   # 12 dígitos: 1º descartado
        ("1234567890", None),              # 10 dígitos: rejeitado
        ("1234567890123", None),           # 13 dígitos: rejeitado
        ("", None),
        (None, None),
    ],
)
def test_pis_regras_11_e_12_digitos(valor, normalizado):
    assert main.limpar_pis(valor) == normalizado


# ------------------------------------------------------------------- MATRÍCULA

@pytest.mark.parametrize(
    "valor, esperado",
    [
        (1000001, "01000001"),
        ("1000001", "01000001"),
        ("1000001.0", "01000001"),
        ("01000001", "01000001"),
        (None, ""),
    ],
)
def test_matricula_preserva_zeros_a_esquerda(valor, esperado):
    assert main.normalizar_matricula(valor) == esperado


def test_vinculo_pelo_prefixo():
    assert main.vinculo_por_matricula("01000001") == "ENTITY A"
    assert main.vinculo_por_matricula("02000001") == "ENTITY B"
    assert main.vinculo_por_matricula("04000001") == "ENTITY C"
    assert main.vinculo_por_matricula("77000001") == "NÃO IDENTIFICADO"


# ----------------------------------------------------------------------- DATAS

def test_parse_data_formatos_e_sentinelas():
    assert main.parse_data("2021-10-15") == date(2021, 10, 15)
    assert main.parse_data("15/10/2021") == date(2021, 10, 15)
    assert main.parse_data("ATIVO") is None
    assert main.parse_data("") is None
    assert main.parse_data(None) is None
    assert main.eh_ativo("ATIVO") is True
    assert main.eh_ativo("15/10/2021") is False


def test_data_base_e_ultimo_dia_do_mes():
    assert main.ultimo_dia_do_mes(2, 2024) == date(2024, 2, 29)
    assert main.ultimo_dia_do_mes(10, 2026) == date(2026, 10, 31)


# ---------------------------------------------------- MESES COMPLETOS (sem dias)

def test_meses_ignoram_dia_mas_respeitam_mes_e_ano():
    # exemplos do docstring de meses_intervalo
    assert main.meses_intervalo(date(2012, 6, 15), date(2026, 6, 14)) == 168
    assert main.meses_intervalo(date(2021, 11, 1), date(2026, 10, 30)) == 59
    assert main.meses_intervalo(date(2021, 11, 1), date(2026, 11, 30)) == 60


def test_intervalo_invalido_vale_zero():
    assert main.meses_intervalo(date(2026, 1, 1), date(2026, 1, 1)) == 0
    assert main.meses_intervalo(date(2026, 2, 1), date(2026, 1, 1)) == 0


def test_formatar_meses():
    assert main.formatar_meses(59) == "4 anos e 11 meses"
    assert main.formatar_meses(60) == "5 anos e 0 meses"


# ------------------------------------------------------------ MESCLAR INTERVALOS

def iv(inicio, fim=None):
    return {"inicio": inicio, "fim": fim, "aberto": fim is None}


def test_sobreposicao_nao_conta_em_dobro():
    mesclados = main.mesclar_intervalos(
        [
            iv(date(2016, 10, 10)),                    # ativo
            iv(date(2018, 3, 1), date(2020, 8, 31)),   # contido no anterior
        ]
    )
    assert len(mesclados) == 1
    assert main.calcular_meses_totais(mesclados, date(2026, 10, 31)) == 120


def test_continuidade_no_mesmo_dia_nao_gera_hiato():
    mesclados = main.mesclar_intervalos(
        [
            iv(date(2011, 10, 3), date(2017, 6, 20)),
            iv(date(2017, 6, 20)),
        ]
    )
    assert len(mesclados) == 1
    assert mesclados[0]["aberto"] is True


def test_hiato_separa_intervalos_e_nao_conta():
    mesclados = main.mesclar_intervalos(
        [
            iv(date(2010, 1, 10), date(2014, 12, 15)),
            iv(date(2016, 9, 5)),
        ]
    )
    assert len(mesclados) == 2
    # 59 meses (2010-01 -> 2014-12) + 121 meses (2016-09 -> 2026-10)
    assert main.calcular_meses_totais(mesclados, date(2026, 10, 31)) == 180


# ---------------------------------------------------------- MARCOS DE QUINQUÊNIO

def test_marco_de_5_anos_no_mes_correto():
    marcos = main.calcular_datas_quinquenio(
        [iv(date(2021, 11, 1))], date(2026, 11, 30)
    )
    assert [(m, d.year, d.month) for m, d in marcos] == [(5, 2026, 11)]


def test_59_meses_nao_atinge_quinquenio():
    marcos = main.calcular_datas_quinquenio(
        [iv(date(2021, 11, 1))], date(2026, 10, 31)
    )
    assert marcos == []


def test_data_base_e_limite_rigido():
    # admitido depois da data-base: nada é calculado
    marcos = main.calcular_datas_quinquenio(
        [iv(date(2026, 11, 15))], date(2026, 10, 31)
    )
    assert marcos == []


def test_matricula_principal_e_a_ativa_mais_nova():
    regs = [
        {"aberto": False, "inicio": date(2020, 1, 1), "id": "a"},
        {"aberto": True, "inicio": date(2015, 1, 1), "id": "b"},
        {"aberto": True, "inicio": date(2018, 1, 1), "id": "c"},
    ]
    assert main.selecionar_matricula_principal(regs)["id"] == "c"
    sem_ativa = [{**r, "aberto": False} for r in regs]
    assert main.selecionar_matricula_principal(sem_ativa)["id"] == "a"
