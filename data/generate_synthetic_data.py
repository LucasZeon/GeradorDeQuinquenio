"""
Gera o banco SQLite SINTÉTICO usado pela versão pública do projeto.

Tudo aqui é fictício: nomes ("Employee 001"), matrículas (prefixo + sequência),
CPF/PIS (strings de 11 dígitos sem relação com pessoas reais), departamentos,
cargos e entidades. A geração é determinística (seed fixa).

O esquema (employee / department / employee_attribute) é fictício e imita
apenas a *forma* do problema original: várias matrículas por pessoa,
atributos em tabela chave-valor, data sentinela para "vínculo ativo".

Uso:
    python data/generate_synthetic_data.py [caminho_do_banco.db]

Cenários nomeados (CASES) cobrem as regras do cálculo para a data-base
10/2026 (último dia = 31/10/2026). O dicionário EXPECTED descreve o resultado
esperado por modo de agrupamento e é usado pelos testes.
"""

import os
import random
import sqlite3
import sys
from datetime import date

SEED = 42
ACTIVE_SENTINEL = "1899-12-30"  # data sentinela: "sem rescisão"

DEFAULT_DB = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "synthetic_hr.db",
)

# prefixo da matrícula -> código da entidade empregadora (fictícia)
EMPLOYER_BY_PREFIX = {
    "01": "ENT-A",
    "02": "ENT-B",
    "04": "ENT-C1",
}

SCHEMA = """
CREATE TABLE department (
    department_id TEXT NOT NULL,
    company_code  TEXT NOT NULL,
    description   TEXT NOT NULL,
    PRIMARY KEY (department_id, company_code)
);

CREATE TABLE employee (
    badge_id         TEXT NOT NULL,   -- "matrícula" (8 posições, texto)
    company_code     TEXT NOT NULL,
    department_id    TEXT NOT NULL,
    full_name        TEXT NOT NULL,
    pis              TEXT,
    daily_minutes    INTEGER,
    hire_date        TEXT,            -- ISO yyyy-mm-dd
    termination_date TEXT,            -- ISO; '1899-12-30' = ativo
    employer_code    TEXT,
    PRIMARY KEY (badge_id, company_code)
);

CREATE TABLE employee_attribute (
    badge_id       TEXT NOT NULL,
    company_code   TEXT NOT NULL,
    attribute_code TEXT NOT NULL,     -- 'TAX_ID' | 'JOB_TITLE'
    value          TEXT,
    PRIMARY KEY (badge_id, company_code, attribute_code)
);
"""


def d(y, m, day):
    return date(y, m, day).isoformat()


def fake_tax_id(n):
    """11 dígitos sintéticos (sem relação com registros reais)."""
    return f"999{n:06d}00"


def fake_pis(n):
    """11 dígitos sintéticos (sem relação com registros reais)."""
    return f"888{n:07d}0"


class Builder:
    def __init__(self):
        self.persons = []          # metadados por pessoa
        self.rows = []             # linhas de employee
        self.attrs = []            # linhas de employee_attribute
        self._seq = {"01": 0, "02": 0, "04": 0, "99": 0}
        self._n = 0

    def _badge(self, prefix):
        self._seq[prefix] += 1
        return f"{prefix}{self._seq[prefix]:06d}"

    def add_person(
        self,
        name,
        contracts,
        dept="01",
        job="Job Title A",
        minutes=480,
        tax_id="auto",
        pis="auto",
        pis_by_contract=None,
    ):
        """
        contracts: lista de (prefixo, admissão_iso, rescisão_iso|None,
                             company_code[, employer_code])
        tax_id / pis: "auto" gera valores sintéticos; None grava vazio/NULL;
        qualquer outra string é gravada como está (para casos inválidos).
        """
        self._n += 1
        n = self._n
        cpf = fake_tax_id(n) if tax_id == "auto" else tax_id
        pis_v = fake_pis(n) if pis == "auto" else pis
        badges = []

        for i, c in enumerate(contracts):
            prefix, hire, term, company = c[:4]
            employer = (
                c[4] if len(c) > 4 else EMPLOYER_BY_PREFIX.get(prefix)
            )
            badge = self._badge(prefix)
            badges.append(badge)
            pis_row = pis_v
            if pis_by_contract:
                pis_row = pis_by_contract[i]
            self.rows.append(
                (
                    badge,
                    company,
                    dept,
                    name,
                    pis_row,
                    minutes,
                    hire,
                    term if term else ACTIVE_SENTINEL,
                    employer,
                )
            )
            self.attrs.append((badge, company, "TAX_ID", cpf))
            self.attrs.append((badge, company, "JOB_TITLE", job))

        self.persons.append({"name": name, "badges": badges})
        return badges


def build_cases(b):
    """Cenários nomeados. Data-base do cálculo: 31/10/2026."""

    # 001 — uma matrícula ativa; completa exatamente 5 anos em 10/2026.
    b.add_person(
        "Employee 001",
        [("01", d(2021, 10, 15), None, "CO1")],
        dept="01",
        job="Job Title A",
    )

    # 002 — admitido em 01/11/2021: 59 meses até 31/10/2026 (ainda não 60).
    b.add_person(
        "Employee 002",
        [("01", d(2021, 11, 1), None, "CO1")],
        dept="02",
        job="Job Title B",
    )

    # 003 — continuidade: 2ª matrícula começa NO MESMO DIA da rescisão da 1ª
    # (sem hiato). Tempo contínuo desde 2011 -> 15 anos em 10/2026.
    b.add_person(
        "Employee 003",
        [
            ("01", d(2011, 10, 3), d(2017, 6, 20), "CO1"),
            ("01", d(2017, 6, 20), None, "CO1"),
        ],
        dept="03",
        job="Job Title C",
    )

    # 004 — HIATO entre vínculos (o intervalo sem vínculo não conta).
    # 59 meses + 121 meses = 180 -> 15 anos em 10/2026. Sem o hiato ser
    # descontado, o marco teria ocorrido muito antes.
    b.add_person(
        "Employee 004",
        [
            ("01", d(2010, 1, 10), d(2014, 12, 15), "CO1"),
            ("01", d(2016, 9, 5), None, "CO1"),
        ],
        dept="01",
        job="Job Title A",
    )

    # 005 — matrículas SIMULTÂNEAS: a 2ª está totalmente contida na 1ª e não
    # pode contar tempo em dobro. 120 meses -> 10 anos em 10/2026.
    b.add_person(
        "Employee 005",
        [
            ("01", d(2016, 10, 10), None, "CO1"),
            ("02", d(2018, 3, 1), d(2020, 8, 31), "CO1"),
        ],
        dept="04",
        job="Job Title D",
    )

    # 006 — vínculo encerrado no próprio mês do marco (5 anos em 20/10/2026).
    b.add_person(
        "Employee 006",
        [("02", d(2021, 10, 20), d(2026, 10, 20), "CO1")],
        dept="02",
        job="Job Title B",
    )

    # 007 — rescisão é data EXCLUSIVA / mês incompleto: 59 meses, não conta.
    b.add_person(
        "Employee 007",
        [("02", d(2021, 10, 1), d(2026, 9, 30), "CO2")],
        dept="05",
        job="Job Title E",
    )

    # 008 — três matrículas em entidades diferentes (prefixos 01, 02 e 04),
    # encadeadas no mesmo dia. 15 anos em 10/2026; principal = a ativa.
    b.add_person(
        "Employee 008",
        [
            ("01", d(2011, 10, 10), d(2016, 2, 1), "CO1"),
            ("02", d(2016, 2, 1), d(2021, 3, 15), "CO2"),
            ("04", d(2021, 3, 15), None, "CO1"),
        ],
        dept="03",
        job="Job Title C",
    )

    # 009 — PIS com 12 dígitos numa matrícula (1º dígito descartado) e 11 na
    # outra: só agrupam corretamente no modo PIS graças a essa regra.
    pis9 = fake_pis(9)
    b.add_person(
        "Employee 009",
        [
            ("01", d(2016, 10, 5), d(2019, 1, 10), "CO1"),
            ("01", d(2019, 1, 10), None, "CO1"),
        ],
        dept="02",
        job="Job Title B",
        pis=pis9,
        pis_by_contract=["7" + pis9, pis9],
    )

    # 010 — CPF com 10 dígitos: inválido no modo CPF; válido no modo PIS.
    b.add_person(
        "Employee 010",
        [("01", d(2021, 10, 12), None, "CO1")],
        dept="01",
        job="Job Title A",
        tax_id="9990000100",
    )

    # 011 — CPF vazio: inválido no modo CPF; válido no modo PIS.
    b.add_person(
        "Employee 011",
        [("01", d(2016, 10, 21), None, "CO1")],
        dept="03",
        job="Job Title C",
        tax_id="",
    )

    # 012 — PIS com 10 dígitos: inválido no modo PIS; válido no modo CPF.
    b.add_person(
        "Employee 012",
        [("01", d(2011, 10, 8), None, "CO1")],
        dept="04",
        job="Job Title D",
        pis="8880000120",
    )

    # 013 — PIS ausente (NULL): inválido no modo PIS; válido no modo CPF.
    b.add_person(
        "Employee 013",
        [("01", d(2006, 10, 2), None, "CO1")],
        dept="05",
        job="Job Title E",
        pis=None,
    )

    # 014 — admissão vazia: inválido em ambos os modos.
    b.add_person(
        "Employee 014",
        [("01", "", None, "CO1")],
        dept="01",
        job="Job Title A",
    )

    # 015 — admissão posterior à data-base: ignorado pelo cálculo.
    b.add_person(
        "Employee 015",
        [("01", d(2026, 11, 15), None, "CO1")],
        dept="02",
        job="Job Title B",
    )

    # Registros que a própria consulta SQL deve excluir:
    #  - matrículas técnicas reservadas (99999901 / 99999902);
    #  - empresa CO3 (fora do filtro).
    for badge, name in (
        ("99999901", "Employee 901 (technical)"),
        ("99999902", "Employee 902 (technical)"),
    ):
        b.rows.append(
            (badge, "CO1", "01", name, fake_pis(901), 480,
             d(2021, 10, 1), ACTIVE_SENTINEL, "ENT-X")
        )
        b.attrs.append((badge, "CO1", "TAX_ID", fake_tax_id(901)))
        b.attrs.append((badge, "CO1", "JOB_TITLE", "Job Title A"))

    b.add_person(
        "Employee 903 (other company)",
        [("01", d(2021, 10, 1), None, "CO3")],
        dept="01",
    )


def build_filler(b, rnd, quantity=30):
    """
    Massa de fundo: vínculos ativos com admissão em meses != 10, para nunca
    coincidirem com o marco de 10/2026 e não afetarem as verificações.
    """
    months = [1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12]
    prefixes = ["01", "02", "04"]
    jobs = [f"Job Title {c}" for c in "ABCDE"]
    for i in range(quantity):
        prefix = rnd.choice(prefixes)
        hire = d(rnd.randint(2001, 2025), rnd.choice(months), rnd.randint(1, 28))
        employer = EMPLOYER_BY_PREFIX[prefix]
        if prefix == "04":
            employer = rnd.choice(["ENT-C1", "ENT-C2", "ENT-X"])
        b.add_person(
            f"Employee {101 + i}",
            [(prefix, hire, None, rnd.choice(["CO1", "CO2"]), employer)],
            dept=f"{rnd.randint(1, 5):02d}",
            job=rnd.choice(jobs),
            minutes=rnd.choice([360, 420, 480]),
        )


# Resultado esperado para 10/2026 (matrícula "principal" omitida).
# Valor = quinquênio (em anos) atingido no mês; ausente = não selecionado.
EXPECTED = {
    "CPF": {
        "selected": {
            "Employee 001": 5,
            "Employee 003": 15,
            "Employee 004": 15,
            "Employee 005": 10,
            "Employee 006": 5,
            "Employee 008": 15,
            "Employee 009": 10,
            "Employee 012": 15,
            "Employee 013": 20,
        },
        "invalid_names": {"Employee 010", "Employee 011", "Employee 014"},
    },
    "PIS": {
        "selected": {
            "Employee 001": 5,
            "Employee 003": 15,
            "Employee 004": 15,
            "Employee 005": 10,
            "Employee 006": 5,
            "Employee 008": 15,
            "Employee 009": 10,
            "Employee 010": 5,
            "Employee 011": 10,
        },
        "invalid_names": {"Employee 012", "Employee 013", "Employee 014"},
    },
}


def build_database(path=DEFAULT_DB):
    rnd = random.Random(SEED)
    b = Builder()
    build_cases(b)
    build_filler(b, rnd)

    if os.path.exists(path):
        os.remove(path)
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)

    conn = sqlite3.connect(path)
    try:
        conn.executescript(SCHEMA)
        for company in ("CO1", "CO2", "CO3"):
            for i in range(1, 6):
                conn.execute(
                    "INSERT INTO department VALUES (?, ?, ?)",
                    (f"{i:02d}", company, f"Department {i:02d}"),
                )
        conn.executemany(
            "INSERT INTO employee VALUES (?,?,?,?,?,?,?,?,?)", b.rows
        )
        conn.executemany(
            "INSERT INTO employee_attribute VALUES (?,?,?,?)", b.attrs
        )
        conn.commit()
    finally:
        conn.close()
    return path, len(b.rows)


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_DB
    caminho, total = build_database(target)
    print(f"Banco sintético gerado: {caminho} ({total} matrículas)")
