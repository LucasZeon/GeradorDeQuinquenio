"""
Testes da verificação de segurança e do .gitignore de lista de permissão.

Os "iscas" (dados de mentira que deveriam ser barrados) são montados em tempo
de execução, por concatenação, para que este arquivo de teste não contenha,
ele mesmo, nenhum padrão que a verificação bloqueia.
"""

import os
import shutil
import struct
import subprocess
import zlib

import pytest

import check_public_safety as safety

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def labels(text, terms=(), secrets=()):
    return {label for _, label in safety.check_content("x.py", text, list(terms), list(secrets))}


# ------------------------------------------------------------- caminhos/arquivos

@pytest.mark.parametrize(
    "path",
    [
        ".env.local",
        "chave.pfx",
        "dados/funcionarios.csv",
        "planilha.xlsx",
        "app.exe",
        "lib/oci.dll",
        "config.ini",
        "tns" + "names.ora",
        "usuarios_config.enc.json",
        "credentials.yaml",
        "build/x.py",
        "src/__pycache__/main.cpython-314.pyc",
        ".safety-terms.local",
        "img/foto.png",
        "docs/foto.png",
        "docs/img/dados.csv",
    ],
)
def test_caminhos_proibidos_sao_barrados(path):
    assert safety.check_path(path)


@pytest.mark.parametrize(
    "path",
    [
        "src/main.py",
        "README.md",
        "sql/quintennial_query.sql",
        "data/generate_synthetic_data.py",
        ".gitignore",
        ".githooks/pre-commit",
        ".github/workflows/ci.yml",
        ".github/dependabot.yml",
        "SECURITY.md",
        "docs/img/gui.png",
        "requirements.txt",
    ],
)
def test_caminhos_legitimos_passam(path):
    assert safety.check_path(path) == []


# ------------------------------------------------------------------- conteúdo

def test_atribuicao_de_credencial():
    assert "atribuição de credencial" in labels("pass" + "word = 'abc12345'")
    assert "atribuição de credencial" in labels("SENHA_" + "X = \"abc12345\"")


def test_placeholders_nao_sao_credencial():
    assert "atribuição de credencial" not in labels("pass" + "word = '[REDACTED]'")
    assert "atribuição de credencial" not in labels("token = 'your-token-here'")


def test_caminho_pessoal_windows():
    line = "C:" + "\\Users\\" + "alguem\\Desktop"
    assert "caminho pessoal do Windows" in labels(line)


def test_documentos_e_rede():
    cpf = "123." + "456." + "789-" + "09"
    assert "CPF formatado" in labels(cpf)
    ip = ".".join(["10", "20", "30", "40"])
    assert "endereço IP" in labels(ip)
    assert "endereço IP" not in labels("host = " + "127" + ".0.0.1")
    mail = "fulano" + "@" + "empresa-real.com.br"
    assert "e-mail" in labels(mail)
    assert "e-mail" not in labels("fulano" + "@" + "example.com")
    unc = "\\\\" + "servidor01\\pasta"
    assert "caminho de rede (UNC)" in labels(unc)


def test_tokens_e_chaves():
    assert "bloco de chave privada" in labels("-----BEGIN " + "RSA PRIVATE KEY-----")
    jwt = ".".join(["eyJ" + "hbGciOiJIUzI1NiJ9", "eyJzdWIiOiIxMjM0NTY3ODkw", "abcdefghij"])
    assert "JWT" in labels(jwt)
    fernet = "aB3dE6gH9jK2mN5p" + "Q8sT1vW4yZ7bC0eF" + "3hJ6kM9nP2q" + "="
    assert "chave Fernet/base64 de 32 bytes" in labels("k = '" + fernet + "'")
    url = "postgres" + "://" + "usuario:segredo" + "@" + "host/db"
    assert "credencial em URL" in labels(url)
    assert "string de conexão" in labels("jdbc" + ":oracle:thin:@host:1521:x")


def test_string_longa_aleatoria_mas_nao_separadores_ou_nomes():
    aleatoria = "q8Zr2LmXv9TbN4cY" + "h7WsKd1PgJf6UaEo" + "Rt3VxCiBn5MyHw0"
    assert "string longa de aparência aleatória" in labels("x = '" + aleatoria + "'")
    assert not labels("# " + "=" * 75)
    assert not labels("data/generate_synthetic_data.py")
    assert not labels("test_pessoa_com_tres_matriculas_usa_a_ativa_como_principal")


def test_termos_locais_e_segredos_da_pasta_privada():
    assert "termo sensível da lista local" in labels("Empresa Interna S.A.", terms=["empresa interna"])
    assert "valor idêntico a um segredo da pasta privada" in labels(
        "valor = 'meuvalorsecreto'", secrets=["meuvalorsecreto"]
    )


def test_valores_nunca_aparecem_na_saida(tmp_path, capsys):
    arquivo = tmp_path / "x.py"
    arquivo.write_text("pass" + "word = 'valorMuitoSecreto99'\n", encoding="utf-8")
    resultado = safety.check_content("x.py", arquivo.read_text(encoding="utf-8"), [], [])
    assert resultado and "valorMuitoSecreto99" not in repr(resultado)


@pytest.mark.skipif(not safety.in_git_repo(), reason="requer repositório Git")
def test_arvore_atual_do_projeto_passa():
    checked, violations, _, _ = safety.run("all", None, None)
    assert checked > 0
    assert violations == []


# ---------------------------------------------- .gitignore de lista de permissão

@pytest.mark.skipif(shutil.which("git") is None, reason="git não encontrado")
def test_gitignore_barra_arquivos_novos(tmp_path):
    shutil.copy(os.path.join(ROOT, ".gitignore"), tmp_path / ".gitignore")
    subprocess.run(["git", "init", "-q", str(tmp_path)], check=True)

    iscas = [
        ".env", ".env.local", ".env.production", "credentials.json", "secrets.yaml",
        "token.txt", "config.ini", "settings.yaml", "id_rsa", "cert.pfx", "chave.pem",
        "a.dll", "a.exe", "dados.csv", "real.xls", "real.xlsm", "doc.pdf", "ata.docx",
        "foto.png", "dump.sql.gz", "backup.zip", "tns" + "names.ora", "x.db", "x.sqlite",
        "output/relatorio.xlsx", "src/novo_dado.csv", "src/x.env", "docs/img.png",
        "tests/dados.xlsx", "data/real.db", "docs/img/foto.jpg", "docs/img/dados.csv", "oracle_config/x", "build/x", "dist/x",
        "src/__pycache__/m.pyc", "notas.txt", "desktop.ini", ".DS_Store",
    ]
    for isca in iscas:
        alvo = tmp_path / isca
        alvo.parent.mkdir(parents=True, exist_ok=True)
        alvo.write_text("x", encoding="utf-8")

    for nome in ["README.md", "src/main.py", "tests/t.py", "sql/q.sql",
                 "docs/d.md", "tools/check_public_safety.py",
                 ".githooks/pre-commit", ".github/workflows/ci.yml",
                 ".github/dependabot.yml", "SECURITY.md", "docs/img/gui.png"]:
        alvo = tmp_path / nome
        alvo.parent.mkdir(parents=True, exist_ok=True)
        alvo.write_text("x", encoding="utf-8")

    saida = subprocess.run(
        ["git", "ls-files", "--others", "--exclude-standard"],
        cwd=tmp_path, capture_output=True, text=True, check=True,
    ).stdout.split()

    esperados = {".gitignore", "README.md", "src/main.py", "tests/t.py", "sql/q.sql",
                 "docs/d.md", "tools/check_public_safety.py", ".githooks/pre-commit",
                 ".github/workflows/ci.yml", ".github/dependabot.yml", "SECURITY.md",
                 "docs/img/gui.png"}
    assert set(saida) == esperados


# ------------------------------------------------------------------ imagens PNG

def _png(*extra_chunks):
    def chunk(kind, data):
        body = kind + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body))

    ihdr = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    raw = safety.PNG_SIGNATURE + chunk(b"IHDR", ihdr)
    for kind, data in extra_chunks:
        raw += chunk(kind, data)
    return raw + chunk(b"IDAT", zlib.compress(b"\x00\x00\x00\x00")) + chunk(b"IEND", b"")


def test_png_limpo_passa():
    assert safety.check_png(_png()) == []


@pytest.mark.parametrize("kind", [b"tEXt", b"iTXt", b"zTXt", b"eXIf"])
def test_png_com_metadados_e_barrado(kind):
    problemas = safety.check_png(_png((kind, b"Author\x00alguem")))
    assert problemas and kind.decode() in problemas[0]


def test_arquivo_que_nao_e_png_e_barrado():
    assert safety.check_png(b"nao sou uma imagem")
