"""
Organizador de Contracheques
-----------------------------
Consolida múltiplos PDFs de contracheques em um único arquivo,
ordenados cronologicamente (ordem crescente por mês/ano).

Uso:
    python organizar_contracheques.py <PASTA_ENTRADA> [ARQUIVO_SAIDA]

Exemplos:
    python organizar_contracheques.py ./contracheques
    python organizar_contracheques.py ./contracheques ./resultado.pdf

Dependências:
    pip install pdfplumber pypdf
"""

import os
import re
import sys
from datetime import datetime

import pdfplumber
from pypdf import PdfReader, PdfWriter


# Padrões comuns de mês/ano encontrados em contracheques
# Exemplos: "12/2023", "06/2020", "02.2026", "MES/ANO: 12/2023"
PADRAO_MES_ANO = re.compile(
    r'(?:M[eê]s\s*/?\s*Ano\s*[:\s]*)?'
    r'(\d{1,2})\s*[/.\-]\s*(2\d{3})',
    re.IGNORECASE,
)

# Padrão para datas por extenso: "Janeiro de 2023", "FEVEREIRO/2024"
MESES_EXTENSO = {
    'janeiro': 1, 'fevereiro': 2, 'março': 3, 'marco': 3,
    'abril': 4, 'maio': 5, 'junho': 6,
    'julho': 7, 'agosto': 8, 'setembro': 9,
    'outubro': 10, 'novembro': 11, 'dezembro': 12,
}

PADRAO_MES_EXTENSO = re.compile(
    r'(' + '|'.join(MESES_EXTENSO.keys()) + r')\s*[/\s]+de?\s*(2\d{3})',
    re.IGNORECASE,
)

# Padrão "Competência" ou "Referência" seguido de mês/ano
PADRAO_COMPETENCIA = re.compile(
    r'(?:compet[eê]ncia|refer[eê]ncia)\s*[:\s]*(\d{1,2})\s*[/.\-]\s*(2\d{3})',
    re.IGNORECASE,
)


def extrair_data_contracheque(caminho_pdf: str) -> datetime | None:
    """
    Extrai a data (mês/ano) de um PDF de contracheque.

    Tenta múltiplos padrões para ser compatível com diferentes
    formatos de contracheque (PREVI, INSS, empresas diversas).
    """
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            # Lê apenas a primeira página (onde fica o cabeçalho)
            if not pdf.pages:
                return None
            texto = pdf.pages[0].extract_text() or ""
    except Exception as e:
        print(f"  [ERRO] Não foi possível ler '{caminho_pdf}': {e}")
        return None

    # Tenta padrão de competência/referência primeiro (mais específico)
    match = PADRAO_COMPETENCIA.search(texto)
    if match:
        mes, ano = int(match.group(1)), int(match.group(2))
        if 1 <= mes <= 12:
            return datetime(ano, mes, 1)

    # Tenta mês por extenso
    match = PADRAO_MES_EXTENSO.search(texto)
    if match:
        nome_mes = match.group(1).lower()
        mes = MESES_EXTENSO.get(nome_mes)
        ano = int(match.group(2))
        if mes:
            return datetime(ano, mes, 1)

    # Tenta padrão numérico MM/AAAA (pega a primeira ocorrência relevante)
    # Filtra datas que parecem ser cabeçalho do contracheque
    for match in PADRAO_MES_ANO.finditer(texto):
        mes, ano = int(match.group(1)), int(match.group(2))
        if 1 <= mes <= 12 and 2000 <= ano <= 2099:
            return datetime(ano, mes, 1)

    return None


def organizar_contracheques(pasta_entrada: str, arquivo_saida: str | None = None) -> str:
    """
    Lê todos os PDFs da pasta, extrai as datas, ordena cronologicamente
    e gera um único PDF compilado.

    Retorna o caminho do arquivo de saída.
    """
    if not os.path.isdir(pasta_entrada):
        print(f"ERRO: Pasta '{pasta_entrada}' não encontrada.")
        sys.exit(1)

    # Listar todos os PDFs da pasta
    arquivos_pdf = [
        os.path.join(pasta_entrada, f)
        for f in os.listdir(pasta_entrada)
        if f.lower().endswith('.pdf')
    ]

    if not arquivos_pdf:
        print(f"ERRO: Nenhum arquivo PDF encontrado em '{pasta_entrada}'.")
        sys.exit(1)

    print(f"Encontrados {len(arquivos_pdf)} arquivo(s) PDF.\n")

    # Extrair datas e associar aos arquivos
    contracheques = []
    sem_data = []

    for caminho in arquivos_pdf:
        nome = os.path.basename(caminho)
        data = extrair_data_contracheque(caminho)
        if data:
            contracheques.append((data, caminho))
            print(f"  OK  {nome} -> {data.strftime('%m/%Y')}")
        else:
            sem_data.append(caminho)
            print(f"  ??  {nome} -> Data não encontrada")

    if sem_data:
        print(f"\nATENÇÃO: {len(sem_data)} arquivo(s) sem data identificada:")
        for c in sem_data:
            print(f"  - {os.path.basename(c)}")
        print("Esses arquivos serão adicionados ao final do compilado.\n")

    if not contracheques and not sem_data:
        print("Nenhum contracheque processado.")
        sys.exit(1)

    # Verificar duplicatas
    datas_vistas = {}
    for data, caminho in contracheques:
        chave = data.strftime('%m/%Y')
        if chave in datas_vistas:
            print(f"  DUPLICATA: {chave} aparece em:")
            print(f"    - {os.path.basename(datas_vistas[chave])}")
            print(f"    - {os.path.basename(caminho)}")
        datas_vistas[chave] = caminho

    # Ordenar por data crescente
    contracheques.sort(key=lambda x: x[0])

    # Montar PDF compilado
    writer = PdfWriter()

    for data, caminho in contracheques:
        reader = PdfReader(caminho)
        for page in reader.pages:
            writer.add_page(page)

    # Adicionar arquivos sem data ao final
    for caminho in sem_data:
        reader = PdfReader(caminho)
        for page in reader.pages:
            writer.add_page(page)

    # Definir nome do arquivo de saída
    if arquivo_saida is None:
        if contracheques:
            primeira = contracheques[0][0].strftime('%m.%Y')
            ultima = contracheques[-1][0].strftime('%m.%Y')
            nome_saida = f"Contracheques_{primeira}_a_{ultima}.pdf"
        else:
            nome_saida = "Contracheques_compilado.pdf"
        arquivo_saida = os.path.join(pasta_entrada, nome_saida)

    writer.write(arquivo_saida)
    total_paginas = len(writer.pages)

    print(f"\nCompilado gerado com sucesso!")
    print(f"  Arquivo: {arquivo_saida}")
    print(f"  Contracheques: {len(contracheques)} (+ {len(sem_data)} sem data)")
    print(f"  Total de páginas: {total_paginas}")

    if contracheques:
        print(f"  Período: {contracheques[0][0].strftime('%m/%Y')} a "
              f"{contracheques[-1][0].strftime('%m/%Y')}")

    return arquivo_saida


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        print("Uso: python organizar_contracheques.py <PASTA_ENTRADA> [ARQUIVO_SAIDA]")
        sys.exit(1)

    pasta_entrada = sys.argv[1]
    arquivo_saida = sys.argv[2] if len(sys.argv) > 2 else None

    organizar_contracheques(pasta_entrada, arquivo_saida)


if __name__ == "__main__":
    main()
