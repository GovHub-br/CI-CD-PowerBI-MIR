import os
import re
import sys
from pathlib import Path

PBIP_DIR = Path(os.environ.get("PBIP_DIR", "Dashboard"))
SITE_URL = os.environ.get("SHAREPOINT_SITE_URL", "").strip()
FOLDER = os.environ.get("SHAREPOINT_FOLDER", "").strip().strip("/")
MODE = os.environ.get("CONNECTION_MODE", "sharepoint").strip().lower()


def build_web(subfolder: str, filename: str) -> str:
    site = SITE_URL.rstrip("/") + "/"
    parts = [FOLDER, subfolder, filename] if subfolder else [FOLDER, filename]
    rel = "/".join(p for p in parts if p)
    return f'Web.Contents("{site}", [RelativePath="{rel}"])'


def build_sharepoint(subfolder: str, filename: str) -> str:
    site = SITE_URL.rstrip("/") + "/"
    folder_parts = [site.rstrip("/"), FOLDER, subfolder] if subfolder else [site.rstrip("/"), FOLDER]
    folder_url = "/".join(p.strip("/") for p in folder_parts if p) + "/"
    if not folder_url.startswith("http"):
        folder_url = site + folder_url
    return (
        f'SharePoint.Files("{site}", [ApiVersion=15])'
        f'{{[Name="{filename}", #"Folder Path"="{folder_url}"]}}'
        f'[Content]'
    )


def build_expression(subfolder: str, filename: str) -> str:
    if MODE == "sharepoint":
        return build_sharepoint(subfolder, filename)
    return build_web(subfolder, filename)


SOURCE_RE = re.compile(
    r'File\.Contents\("(?P<local>[^"]+)"\)'
    r'|Web\.Contents\("[^"]*",\s*\[RelativePath="(?P<rel>[^"]+)"\]\)'
    r'|SharePoint\.Files\("[^"]*",\s*\[ApiVersion=15\]\)'
    r'\{\[Name="(?P<name>[^"]+)",\s*#"Folder Path"="(?P<spfolder>[^"]*)"\]\}\[Content\]'
)


def extract_subfolder_and_file(match) -> tuple[str, str]:
    """Retorna (subfolder, filename) de qualquer formato de fonte."""
    if match.group("local"):
        p = Path(match.group("local").replace("\\", "/"))
        return p.parent.name, p.name

    if match.group("rel"):
        p = Path(match.group("rel"))
        # parent.name e a subpasta imediatamente acima do arquivo
        return p.parent.name, p.name

    # SharePoint.Files: subfolder extraida do Folder Path
    filename = match.group("name")
    sp = match.group("spfolder").rstrip("/")
    subfolder = sp.split("/")[-1] if sp else ""
    return subfolder, filename


def replace_connection(content: str):
    replaced = []

    def substituir(match):
        subfolder, filename = extract_subfolder_and_file(match)
        replaced.append(f"{subfolder}/{filename}" if subfolder else filename)
        return build_expression(subfolder, filename)

    new_content = SOURCE_RE.sub(substituir, content)
    return new_content, replaced


def main() -> None:
    if not SITE_URL:
        print("ERROR: SHAREPOINT_SITE_URL nao configurado.", file=sys.stderr)
        print("  Settings -> Secrets and variables -> Actions -> Variables -> New repository variable", file=sys.stderr)
        print("  ex: https://sdhgovbr-my.sharepoint.com/personal/sinapir_igualdaderacial_gov_br/", file=sys.stderr)
        sys.exit(1)
    if not FOLDER:
        print("ERROR: SHAREPOINT_FOLDER nao configurado.", file=sys.stderr)
        print("  ex: Documents/DAMGI/Plaforma Maria Firmina/modelagemBI", file=sys.stderr)
        sys.exit(1)
    if MODE not in ("web", "sharepoint"):
        print(f"ERROR: CONNECTION_MODE invalido: '{MODE}' (use 'web' ou 'sharepoint').", file=sys.stderr)
        sys.exit(1)
    if not PBIP_DIR.exists():
        print(f"ERROR: Pasta '{PBIP_DIR}' nao encontrada.", file=sys.stderr)
        sys.exit(1)

    tmdl_files = list(PBIP_DIR.rglob("*.tmdl"))
    if not tmdl_files:
        print(f"Nenhum arquivo .tmdl encontrado em '{PBIP_DIR}'.")
        sys.exit(0)

    print(f"Modo: {MODE} | Site: {SITE_URL} | Pasta base: {FOLDER}")
    print(f"Verificando {len(tmdl_files)} arquivo(s) .tmdl...\n")

    total = []
    for tmdl_file in tmdl_files:
        content = tmdl_file.read_text(encoding="utf-8")
        new_content, replaced = replace_connection(content)
        if replaced:
            tmdl_file.write_text(new_content, encoding="utf-8")
            for name in replaced:
                print(f"  OK  {tmdl_file.name} -> {name}")
            total.extend(replaced)

    if not total:
        print("Nenhuma conexao File.Contents() encontrada para substituir.")
    else:
        print(f"\nConexoes atualizadas: {len(total)}")


if __name__ == "__main__":
    main()
