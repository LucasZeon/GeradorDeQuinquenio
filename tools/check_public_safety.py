"""
Verificação de segurança do repositório público.

Procura, antes de um commit (ou no CI), tudo aquilo que NUNCA deve ser
publicado: tipos de arquivo proibidos, segredos, caminhos pessoais, dados
pessoais, infraestrutura interna e termos sensíveis definidos localmente.

Uso:
    python tools/check_public_safety.py --staged   # arquivos no "git add" (hook)
    python tools/check_public_safety.py --all      # tudo que o Git adicionaria
    python tools/check_public_safety.py --paths a.py b.md

Sai com código 1 se encontrar qualquer problema. O relatório mostra arquivo,
linha e o tipo da regra, mas NUNCA o valor encontrado.

Camadas opcionais de comparação (não ficam no repositório):

  1) Arquivo local ".safety-terms.local" na raiz, um termo por linha
     (nomes de empresas, de sistemas, de usuário etc.). Linhas com "#" são
     comentários. O arquivo é ignorado pelo Git de propósito: colocar a lista
     no repositório público revelaria justamente os termos.

  2) Pasta de origem privada: informe --original-dir CAMINHO, a variável de
     ambiente SAFETY_ORIGINAL_DIR ou a configuração local do Git
         git config --local safety.originalDir CAMINHO
     O script lê essa pasta, extrai (em memória) os valores atribuídos a
     variáveis do tipo senha/usuário/DSN/token/chave e falha se algum deles
     aparecer nos arquivos públicos. Os valores nunca são impressos.
"""

import argparse
import math
import os
import re
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
THIS_FILE = os.path.normpath("tools/check_public_safety.py")

# Arquivos de texto aceitos no repositório (o resto precisa ser liberado aqui).
ALLOWED_EXTENSIONS = {".py", ".md", ".sql", ".txt", ".yml", ".yaml"}
ALLOWED_BASENAMES = {".gitignore", ".gitattributes", "LICENSE", "pre-commit"}
MAX_BYTES = 1_000_000

FORBIDDEN_EXTENSIONS = {
    ".exe", ".dll", ".so", ".dylib", ".msi", ".pyc", ".pyo", ".spec",
    ".xls", ".xlsx", ".xlsm", ".xlsb", ".csv", ".tsv", ".db", ".sqlite",
    ".sqlite3", ".dump", ".dmp", ".bak", ".old", ".orig", ".log", ".tmp",
    ".pdf", ".doc", ".docx", ".ppt", ".pptx", ".png", ".jpg", ".jpeg",
    ".gif", ".bmp", ".zip", ".7z", ".rar", ".tar", ".gz", ".ora", ".enc",
    ".pem", ".key", ".pfx", ".p12", ".jks", ".keystore", ".env", ".ini",
    ".cfg", ".conf", ".toml", ".local",
}

FORBIDDEN_NAME_PATTERNS = [
    (re.compile(r"^\.env(\..*)?$", re.I), "arquivo .env"),
    (re.compile(r"(credential|secret|token|password|senha)", re.I),
     "nome de arquivo sugere segredo"),
    (re.compile(r"^(id_rsa|id_ed25519|sqlnet|wallet)", re.I),
     "arquivo de chave/infraestrutura"),
    (re.compile(r"\.safety-terms", re.I), "lista local de termos sensíveis"),
    (re.compile(r"\.enc\.json$", re.I), "arquivo criptografado"),
]

FORBIDDEN_PATH_PARTS = {
    "__pycache__", "build", "dist", "_internal", "oracle_config",
    "oracle_client", ".venv", "venv", "output",
}

SECRET_NAME = (
    r"[A-Za-z0-9_]*(?:password|passwd|pwd|senha|secret|token|api_?key|"
    r"access_?key|private_?key|dsn|usuario|username)[A-Za-z0-9_]*"
)
PLACEHOLDER = re.compile(
    r"(redacted|example|changeme|your[_-]|<.*>|\{.*\}|xxx|\*\*\*|placeholder|"
    r"dummy|sample|none|null)",
    re.I,
)

CONTENT_RULES = [
    ("atribuição de credencial",
     re.compile(
         r"(?i)\b" + SECRET_NAME + r"\s*[:=]\s*[\"']([^\"'\s]{3,})[\"']"
     )),
    ("bloco de chave privada",
     re.compile(r"-----BEGIN [A-Z ]*PRIVATE KEY-----")),
    ("chave AWS", re.compile(r"\bAKIA[0-9A-Z]{16}\b")),
    ("chave de API do Google", re.compile(r"\bAIza[0-9A-Za-z_\-]{35}\b")),
    ("token do GitHub", re.compile(r"\bgh[pousr]_[A-Za-z0-9]{30,}\b")),
    ("token do Slack", re.compile(r"\bxox[abprs]-[A-Za-z0-9-]{10,}\b")),
    ("JWT", re.compile(
        r"\beyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{5,}")),
    ("cabeçalho Bearer", re.compile(r"(?i)\bbearer\s+[A-Za-z0-9._\-]{20,}")),
    ("chave Fernet/base64 de 32 bytes",
     re.compile(r"(?<![A-Za-z0-9_\-])[A-Za-z0-9_\-]{43}=(?![A-Za-z0-9])")),
    ("credencial em URL",
     re.compile(r"[A-Za-z][A-Za-z0-9+.\-]*://[^/\s:@]+:[^/\s@]+@")),
    ("string de conexão", re.compile(r"(?i)\bjdbc:|\boracle://|\bDSN=")),
    ("caminho pessoal do Windows",
     re.compile(r"(?i)\b[A-Za-z]:\\+Users\\+(?!Public\b|Default\b|<|%|\*)"
                r"[^\\\s\"']+")),
    ("caminho pessoal Unix",
     re.compile(r"(?<![A-Za-z0-9_.])/(?:Users|home)/(?!<|\$|\{|username\b)"
                r"[A-Za-z0-9._\-]+")),
    ("caminho de rede (UNC)",
     re.compile(r"(?<![A-Za-z0-9\\])\\\\[A-Za-z0-9_.\-]+\\[A-Za-z0-9_$.\-]+")),
    ("CPF formatado", re.compile(r"\b\d{3}\.\d{3}\.\d{3}-\d{2}\b")),
    ("CNPJ formatado", re.compile(r"\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b")),
    ("endereço IP",
     re.compile(r"\b(?:(?:25[0-5]|2[0-4]\d|1?\d?\d)\.){3}"
                r"(?:25[0-5]|2[0-4]\d|1?\d?\d)\b")),
    ("endereço MAC",
     re.compile(r"\b(?:[0-9A-Fa-f]{2}[:-]){5}[0-9A-Fa-f]{2}\b")),
    ("e-mail",
     re.compile(r"(?<![A-Za-z0-9._%+\-])[A-Za-z0-9._%+\-]+@"
                r"([A-Za-z0-9\-]+(?:\.[A-Za-z0-9\-]+)+)")),
]
ALLOWED_IPS = {"127.0.0.1", "0.0.0.0", "255.255.255.255"}
ALLOWED_EMAIL_DOMAINS = {"example.com", "example.org", "example.net",
                         "users.noreply.github.com"}

SECRET_ASSIGN_IN_ORIGINAL = re.compile(
    r"(?i)\b" + SECRET_NAME + r"\s*[:=]\s*[\"']([^\"'\n]{4,})[\"']"
)
ORIGINAL_TEXT_EXTENSIONS = {".py", ".sql", ".json", ".ini", ".cfg", ".txt",
                            ".yml", ".yaml", ".toml", ".env"}
ORIGINAL_SKIP_DIRS = {"__pycache__", "build", "dist", "_internal", ".git",
                      "venv", ".venv", "site-packages", "oracle_client"}


# ----------------------------------------------------------------- utilitários

def run_git(args, cwd=ROOT):
    return subprocess.run(
        ["git", *args], cwd=cwd, capture_output=True, check=False,
    )


def in_git_repo():
    result = run_git(["rev-parse", "--is-inside-work-tree"])
    return result.returncode == 0 and result.stdout.strip() == b"true"


def entropy(text):
    counts = {}
    for ch in text:
        counts[ch] = counts.get(ch, 0) + 1
    size = len(text)
    return -sum(c / size * math.log2(c / size) for c in counts.values())


def looks_like_random_secret(token):
    if len(token) < 40:
        return False
    if len(set(token)) < 12:
        return False
    if re.fullmatch(r"[-=_.#*/\\]+", token):
        return False
    if "/" in token or "\\" in token or token.count(".") > 1:
        return False  # caminhos e nomes de módulo
    if token.count("_") >= 3 or token.count("-") >= 3:
        return False  # identificadores legíveis (snake/kebab-case)
    return entropy(token) >= 4.2


def read_local_terms():
    path = os.path.join(ROOT, ".safety-terms.local")
    if not os.path.isfile(path):
        return []
    with open(path, encoding="utf-8", errors="replace") as handle:
        return [
            line.strip()
            for line in handle
            if line.strip() and not line.lstrip().startswith("#")
        ]


def load_original_secrets(original_dir):
    """Extrai, em memória, valores de variáveis sensíveis da pasta privada."""
    found = {}
    for base, dirs, files in os.walk(original_dir):
        dirs[:] = [d for d in dirs if d not in ORIGINAL_SKIP_DIRS]
        for name in files:
            if os.path.splitext(name)[1].lower() not in ORIGINAL_TEXT_EXTENSIONS:
                continue
            full = os.path.join(base, name)
            try:
                if os.path.getsize(full) > 5_000_000:
                    continue
                with open(full, encoding="utf-8", errors="replace") as handle:
                    text = handle.read()
            except OSError:
                continue
            for match in SECRET_ASSIGN_IN_ORIGINAL.finditer(text):
                value = match.group(1).strip()
                if value and not PLACEHOLDER.search(value):
                    found[value.lower()] = True
    return list(found)


# ------------------------------------------------------------------- verificações

def is_doc_image(rel_path):
    """PNG permitido apenas em docs/img/ (capturas de tela sintéticas)."""
    path = rel_path.replace("\\", "/")
    return path.startswith("docs/img/") and path.lower().endswith(".png")


PNG_SIGNATURE = b"\x89PNG\r\n\x1a\n"
PNG_METADATA_CHUNKS = {b"tEXt", b"zTXt", b"iTXt", b"eXIf"}


def check_png(raw):
    """Recusa PNG inválido ou com metadados embutidos (autor, câmera, GPS...)."""
    if not raw.startswith(PNG_SIGNATURE):
        return ["arquivo não é um PNG válido"]
    problems = []
    pos = len(PNG_SIGNATURE)
    while pos + 8 <= len(raw):
        length = int.from_bytes(raw[pos:pos + 4], "big")
        chunk = raw[pos + 4:pos + 8]
        if chunk in PNG_METADATA_CHUNKS:
            problems.append(
                f"metadados dentro da imagem ({chunk.decode('ascii')})"
            )
        pos += 12 + length
        if chunk == b"IEND":
            break
    return problems


def check_path(rel_path):
    problems = []
    rel = os.path.normpath(rel_path)
    if is_doc_image(rel):
        return problems
    parts = set(rel.replace("\\", "/").split("/"))
    base = os.path.basename(rel)
    ext = os.path.splitext(base)[1].lower()

    if parts & FORBIDDEN_PATH_PARTS:
        problems.append("pasta proibida no caminho")
    for pattern, label in FORBIDDEN_NAME_PATTERNS:
        if pattern.search(base):
            problems.append(label)
    if ext in FORBIDDEN_EXTENSIONS:
        problems.append(f"tipo de arquivo proibido ({ext})")
    elif base not in ALLOWED_BASENAMES and ext not in ALLOWED_EXTENSIONS:
        problems.append("extensão fora da lista de permissão")
    return problems


def check_content(rel_path, text, terms, original_secrets):
    """Retorna lista de (linha, descrição). Nunca inclui o valor achado."""
    problems = []
    is_self = os.path.normpath(rel_path) == THIS_FILE

    lines = text.splitlines()
    lowered = [line.lower() for line in lines]

    for number, line in enumerate(lines, start=1):
        if not is_self:
            for label, pattern in CONTENT_RULES:
                match = pattern.search(line)
                if not match:
                    continue
                if label == "atribuição de credencial":
                    if PLACEHOLDER.search(match.group(1)):
                        continue
                if label == "endereço IP" and match.group(0) in ALLOWED_IPS:
                    continue
                if label == "e-mail":
                    if match.group(1).lower() in ALLOWED_EMAIL_DOMAINS:
                        continue
                problems.append((number, label))
            for token in re.findall(r"[A-Za-z0-9+/=_\-]{40,}", line):
                if looks_like_random_secret(token):
                    problems.append((number, "string longa de aparência aleatória"))
                    break

        for term in terms:
            if term.lower() in lowered[number - 1]:
                problems.append((number, "termo sensível da lista local"))
        for value in original_secrets:
            if value in lowered[number - 1]:
                problems.append(
                    (number, "valor idêntico a um segredo da pasta privada")
                )
    return problems


def files_to_check(mode, explicit_paths):
    if explicit_paths:
        return [(p, None) for p in explicit_paths]
    if mode == "staged":
        out = run_git(["diff", "--cached", "--name-only", "--diff-filter=ACMR",
                       "-z"]).stdout
        names = [n for n in out.decode("utf-8", "replace").split("\0") if n]
        return [(n, "staged") for n in names]
    if in_git_repo():
        out = run_git(["ls-files", "--cached", "--others",
                       "--exclude-standard", "-z"]).stdout
        names = [n for n in out.decode("utf-8", "replace").split("\0") if n]
        return [(n, None) for n in names]
    names = []
    for base, dirs, files in os.walk(ROOT):
        dirs[:] = [d for d in dirs if d != ".git"]
        for name in files:
            names.append(os.path.relpath(os.path.join(base, name), ROOT))
    return [(n, None) for n in names]


def read_content(rel_path, source):
    if source == "staged":
        result = run_git(["show", f":{rel_path}"])
        return result.stdout if result.returncode == 0 else b""
    with open(os.path.join(ROOT, rel_path), "rb") as handle:
        return handle.read()


def run(mode, explicit_paths, original_dir):
    terms = read_local_terms()
    original_secrets = load_original_secrets(original_dir) if original_dir else []

    violations = []
    checked = 0
    for rel_path, source in files_to_check(mode, explicit_paths):
        checked += 1
        for problem in check_path(rel_path):
            violations.append((rel_path, 0, problem))
        try:
            raw = read_content(rel_path, source)
        except OSError:
            continue
        if len(raw) > MAX_BYTES:
            violations.append((rel_path, 0, "arquivo grande demais (> 1 MB)"))
            continue
        if is_doc_image(rel_path):
            for label in check_png(raw):
                violations.append((rel_path, 0, label))
            continue
        if b"\0" in raw:
            violations.append((rel_path, 0, "conteúdo binário"))
            continue
        text = raw.decode("utf-8", errors="replace")
        for number, label in check_content(
            rel_path, text, terms, original_secrets
        ):
            violations.append((rel_path, number, label))

    return checked, violations, len(terms), len(original_secrets)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--staged", action="store_true",
                       help="verifica os arquivos no índice do Git")
    group.add_argument("--all", action="store_true",
                       help="verifica tudo que o Git adicionaria")
    group.add_argument("--paths", nargs="+", help="verifica arquivos específicos")
    parser.add_argument("--original-dir",
                        help="pasta privada de origem para comparação")
    args = parser.parse_args(argv)

    original_dir = (
        args.original_dir
        or os.environ.get("SAFETY_ORIGINAL_DIR")
        or run_git(["config", "--local", "--get", "safety.originalDir"])
        .stdout.decode("utf-8", "replace").strip()
        or None
    )
    if original_dir and not os.path.isdir(original_dir):
        print("AVISO: pasta privada de comparação não encontrada; ignorada.")
        original_dir = None

    mode = "staged" if args.staged else "all"
    checked, violations, n_terms, n_secrets = run(
        mode, args.paths, original_dir
    )

    print(
        f"Arquivos verificados: {checked} | termos locais: {n_terms} | "
        f"valores privados comparados: {n_secrets}"
    )
    if not violations:
        print("OK — nenhuma violação encontrada.")
        return 0

    print(f"FALHA — {len(violations)} problema(s) (valores omitidos):")
    for path, line, label in sorted(set(violations)):
        where = f"{path}:{line}" if line else path
        print(f"  {where}  [{label}]")
    print("Corrija ou remova antes de publicar. Nada foi alterado por este script.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
