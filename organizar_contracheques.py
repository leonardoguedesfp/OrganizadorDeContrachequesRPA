"""
Organizador de Contracheques
-----------------------------
Consolida múltiplos PDFs de contracheques em um único arquivo,
ordenados cronologicamente (ordem crescente por mês/ano).

Modo de uso:
    - Clique duas vezes no .exe (ou execute: python organizar_contracheques.py)
    - Selecione a pasta com os contracheques
    - O compilado será salvo na mesma pasta

Dependências (para rodar via Python):
    pip install pdfplumber pypdf openpyxl

Para gerar o .exe:
    pip install pyinstaller
    pyinstaller --onefile --windowed organizar_contracheques.py
"""

import os
import re
import threading
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

import pdfplumber
from pypdf import PdfReader, PdfWriter

# ---------------------------------------------------------------------------
# Cores e identidade visual
# ---------------------------------------------------------------------------
COR_PRIMARIA = "#003e63"
COR_PRIMARIA_HOVER = "#00527a"
COR_TEXTO_CLARO = "#ffffff"
COR_FUNDO = "#f4f6f8"
COR_FUNDO_LOG = "#ffffff"
COR_BORDA = "#c8d0d8"

# ---------------------------------------------------------------------------
# Padrões comuns de mês/ano encontrados em contracheques
# ---------------------------------------------------------------------------
PADRAO_MES_ANO = re.compile(
    r'(?:M[eê]s\s*/?\s*Ano\s*[:\s]*)?'
    r'(\d{1,2})\s*[/.\-]\s*(2\d{3})',
    re.IGNORECASE,
)

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

PADRAO_COMPETENCIA = re.compile(
    r'(?:compet[eê]ncia|refer[eê]ncia)\s*[:\s]*(\d{1,2})\s*[/.\-]\s*(2\d{3})',
    re.IGNORECASE,
)

# Padrão para valores monetários brasileiros (ex: 1.234,56 ou 1234,56 ou 7.339,36-)
PADRAO_VALOR_BR = re.compile(
    r'^(\d{1,3}(?:\.\d{3})*,\d{2})\s*(-)?$'
)

# Padrão para identificar "Mês / Ano" no cabeçalho da página (campo de referência)
PADRAO_MES_ANO_CABECALHO = re.compile(
    r'[Mm][eê]s\s*/?\s*[Aa]no\s*[:\s]*(\d{1,2}\s*/\s*\d{4})',
)

# Padrão para capturar competência em célula de tabela (MM/AAAA)
PADRAO_COMPETENCIA_CELULA = re.compile(r'(\d{1,2}/\d{4})')


def _parse_valor_br(texto: str) -> float | None:
    """
    Converte valor monetário brasileiro para float.
    Formato: ponto como milhar, vírgula como decimal.
    Sinal negativo pode estar no final (ex: 7.339,36-).
    Retorna None se não for um valor válido.
    """
    if not texto:
        return None
    texto = texto.strip()
    # Detectar sinal negativo (início ou final)
    negativo = False
    if texto.startswith('-'):
        negativo = True
        texto = texto[1:].strip()
    if texto.endswith('-'):
        negativo = True
        texto = texto[:-1].strip()
    match = PADRAO_VALOR_BR.match(texto)
    if match:
        if match.group(2) == '-':
            negativo = True
        valor_str = match.group(1).replace('.', '').replace(',', '.')
        try:
            val = float(valor_str)
            return -val if negativo else val
        except ValueError:
            return None
    # Fallback: tentar interpretar diretamente (ex: "1234,56")
    texto_limpo = texto.replace('.', '').replace(',', '.')
    try:
        val = float(texto_limpo)
        return -val if negativo else val
    except ValueError:
        return None


def extrair_data_contracheque(caminho_pdf: str) -> datetime | None:
    """Extrai a data (mês/ano) de um PDF de contracheque."""
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            if not pdf.pages:
                return None
            texto = pdf.pages[0].extract_text() or ""
    except Exception:
        return None

    match = PADRAO_COMPETENCIA.search(texto)
    if match:
        mes, ano = int(match.group(1)), int(match.group(2))
        if 1 <= mes <= 12:
            return datetime(ano, mes, 1)

    match = PADRAO_MES_EXTENSO.search(texto)
    if match:
        nome_mes = match.group(1).lower()
        mes = MESES_EXTENSO.get(nome_mes)
        ano = int(match.group(2))
        if mes:
            return datetime(ano, mes, 1)

    for match in PADRAO_MES_ANO.finditer(texto):
        mes, ano = int(match.group(1)), int(match.group(2))
        if 1 <= mes <= 12 and 2000 <= ano <= 2099:
            return datetime(ano, mes, 1)

    return None


def _encontrar_indice_coluna(cabecalho: list[str | None], *nomes) -> int | None:
    """Encontra o índice de uma coluna pelo nome (case-insensitive, parcial)."""
    if not cabecalho:
        return None
    for idx, celula in enumerate(cabecalho):
        if not celula:
            continue
        celula_lower = celula.strip().lower()
        for nome in nomes:
            if nome.lower() in celula_lower:
                return idx
    return None


def _extrair_mes_ano_pagina(texto_pagina: str) -> str | None:
    """
    Extrai o campo "Mês / Ano" do cabeçalho de uma página de contracheque.
    Retorna como string no formato MM/AAAA. Não converte para datetime.
    """
    # Prioridade 1: campo "Mês / Ano" explícito
    match = PADRAO_MES_ANO_CABECALHO.search(texto_pagina)
    if match:
        raw = match.group(1).replace(' ', '')
        parts = raw.split('/')
        if len(parts) == 2:
            return f"{int(parts[0]):02d}/{parts[1]}"

    # Prioridade 2: Competência/Referência no cabeçalho
    match = PADRAO_COMPETENCIA.search(texto_pagina)
    if match:
        mes, ano = match.group(1), match.group(2)
        return f"{int(mes):02d}/{ano}"

    # Prioridade 3: primeiro padrão MM/AAAA encontrado
    for match in PADRAO_MES_ANO.finditer(texto_pagina):
        mes, ano = int(match.group(1)), int(match.group(2))
        if 1 <= mes <= 12 and 2000 <= ano <= 2099:
            return f"{mes:02d}/{ano}"

    return None


def _extrair_p300_pagina(page, mes_ano_pagina: str | None):
    """
    Extrai todas as ocorrências de P300 de uma página de PDF.
    Retorna lista de dicts: {contracheque, competencia, valor}
    - contracheque: string MM/AAAA do cabeçalho da página
    - competencia: string MM/AAAA da coluna Competência da linha P300
    - valor: float (negativo se estorno)
    """
    resultados = []

    # Tentar extração por tabelas primeiro
    tabelas = page.extract_tables()
    if tabelas:
        for tabela in tabelas:
            if not tabela or len(tabela) < 2:
                continue

            cabecalho = tabela[0]
            idx_verba = _encontrar_indice_coluna(
                cabecalho, 'verba', 'rubrica', 'código', 'codigo', 'cod',
            )
            idx_valor = _encontrar_indice_coluna(
                cabecalho, 'valor', 'vencimento', 'provento', 'total',
            )
            idx_competencia = _encontrar_indice_coluna(
                cabecalho, 'competência', 'competencia', 'compet',
            )
            idx_desc = _encontrar_indice_coluna(
                cabecalho, 'descrição', 'descricao', 'denominação',
                'denominacao', 'nome',
            )

            for linha in tabela[1:]:
                if not linha:
                    continue

                # Verificar se alguma célula contém "P300"
                tem_p300 = False
                for celula in linha:
                    if celula and re.match(
                        r'^P\s*\.?\s*300$', str(celula).strip(), re.IGNORECASE,
                    ):
                        tem_p300 = True
                        break
                if not tem_p300 and idx_desc is not None and idx_desc < len(linha):
                    if linha[idx_desc] and re.search(
                        r'\bP\s*\.?\s*300\b', str(linha[idx_desc]), re.IGNORECASE,
                    ):
                        tem_p300 = True
                if not tem_p300:
                    continue

                # Extrair valor
                valor = None
                if idx_valor is not None and idx_valor < len(linha):
                    valor = _parse_valor_br(str(linha[idx_valor] or ''))
                if valor is None:
                    for celula in reversed(linha):
                        if celula:
                            v = _parse_valor_br(str(celula))
                            if v is not None:
                                valor = v
                                break
                if valor is None:
                    continue

                # Extrair competência da linha
                competencia = None
                if idx_competencia is not None and idx_competencia < len(linha):
                    cel_comp = str(linha[idx_competencia] or '').strip()
                    m = PADRAO_COMPETENCIA_CELULA.search(cel_comp)
                    if m:
                        raw = m.group(1)
                        parts = raw.split('/')
                        competencia = f"{int(parts[0]):02d}/{parts[1]}"

                resultados.append({
                    'contracheque': mes_ano_pagina or 'N/D',
                    'competencia': competencia or mes_ano_pagina or 'N/D',
                    'valor': valor,
                })

    # Fallback: texto linha a linha (PDFs sem tabela detectável)
    if not tabelas:
        texto = page.extract_text() or ""
        for texto_linha in texto.split('\n'):
            if not re.search(r'\bP\s*\.?\s*300\b', texto_linha, re.IGNORECASE):
                continue
            # Extrair valores monetários (incluindo negativo com - no final)
            valores_raw = re.findall(
                r'(\d{1,3}(?:\.\d{3})*,\d{2}\s*-?)', texto_linha,
            )
            if not valores_raw:
                continue
            # O último valor monetário da linha costuma ser o total
            valor = _parse_valor_br(valores_raw[-1])
            if valor is None:
                continue
            # Tentar pegar competência na mesma linha
            competencia = None
            m = PADRAO_COMPETENCIA_CELULA.search(texto_linha)
            if m:
                raw = m.group(1)
                parts = raw.split('/')
                competencia = f"{int(parts[0]):02d}/{parts[1]}"

            resultados.append({
                'contracheque': mes_ano_pagina or 'N/D',
                'competencia': competencia or mes_ano_pagina or 'N/D',
                'valor': valor,
            })

    return resultados


def extrair_p300_pdf(caminho_pdf: str) -> list[dict]:
    """
    Extrai todas as ocorrências da verba P300 de um PDF.
    Processa cada página independentemente (cada página = um contracheque).
    Retorna lista de dicts: {contracheque, competencia, valor, arquivo}
    Todas as datas são strings MM/AAAA — sem conversão para datetime.
    """
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            if not pdf.pages:
                return []

            resultados = []
            nome_arquivo = os.path.basename(caminho_pdf)

            for page in pdf.pages:
                texto = page.extract_text() or ""
                mes_ano = _extrair_mes_ano_pagina(texto)

                ocorrencias = _extrair_p300_pagina(page, mes_ano)
                for oc in ocorrencias:
                    oc['arquivo'] = nome_arquivo
                resultados.extend(ocorrencias)

            return resultados
    except Exception:
        return []


def organizar_contracheques(pasta_entrada: str, callback_log=None):
    """
    Lê todos os PDFs da pasta, extrai as datas, ordena cronologicamente
    e gera um único PDF compilado na mesma pasta.

    Retorna (sucesso: bool, mensagem: str).
    """
    def log(msg):
        if callback_log:
            callback_log(msg)

    if not os.path.isdir(pasta_entrada):
        return False, f"Pasta '{pasta_entrada}' não encontrada."

    arquivos_pdf = [
        os.path.join(pasta_entrada, f)
        for f in os.listdir(pasta_entrada)
        if f.lower().endswith('.pdf')
    ]

    if not arquivos_pdf:
        return False, "Nenhum arquivo PDF encontrado na pasta selecionada."

    log(f"Encontrados {len(arquivos_pdf)} arquivo(s) PDF.\n")

    contracheques = []
    sem_data = []

    for caminho in arquivos_pdf:
        nome = os.path.basename(caminho)
        data = extrair_data_contracheque(caminho)
        if data:
            contracheques.append((data, caminho))
            log(f"  OK   {nome}  ->  {data.strftime('%m/%Y')}")
        else:
            sem_data.append(caminho)
            log(f"  ??   {nome}  ->  Data não encontrada")

    if sem_data:
        log(f"\n{len(sem_data)} arquivo(s) sem data (serão adicionados ao final).")

    if not contracheques and not sem_data:
        return False, "Nenhum contracheque pôde ser processado."

    # Verificar duplicatas
    datas_vistas = {}
    for data, caminho in contracheques:
        chave = data.strftime('%m/%Y')
        if chave in datas_vistas:
            log(f"\n  DUPLICATA: {chave} aparece em:")
            log(f"    - {os.path.basename(datas_vistas[chave])}")
            log(f"    - {os.path.basename(caminho)}")
        datas_vistas[chave] = caminho

    contracheques.sort(key=lambda x: x[0])

    writer = PdfWriter()

    for data, caminho in contracheques:
        reader = PdfReader(caminho)
        for page in reader.pages:
            writer.add_page(page)

    for caminho in sem_data:
        reader = PdfReader(caminho)
        for page in reader.pages:
            writer.add_page(page)

    if contracheques:
        primeira = contracheques[0][0].strftime('%m.%Y')
        ultima = contracheques[-1][0].strftime('%m.%Y')
        nome_saida = f"Contracheques_{primeira}_a_{ultima}.pdf"
    else:
        nome_saida = "Contracheques_compilado.pdf"

    arquivo_saida = os.path.join(pasta_entrada, nome_saida)
    writer.write(arquivo_saida)
    total_paginas = len(writer.pages)

    resumo = (
        f"\nCompilado gerado com sucesso!\n"
        f"  Arquivo: {nome_saida}\n"
        f"  Contracheques: {len(contracheques)} (+ {len(sem_data)} sem data)\n"
        f"  Total de páginas: {total_paginas}"
    )
    if contracheques:
        resumo += (
            f"\n  Período: {contracheques[0][0].strftime('%m/%Y')} "
            f"a {contracheques[-1][0].strftime('%m/%Y')}"
        )

    log(resumo)
    return True, arquivo_saida


def _chave_ordenacao_mes_ano(texto: str) -> tuple[int, int]:
    """Converte 'MM/AAAA' em (AAAA, MM) para ordenação cronológica."""
    try:
        parts = texto.split('/')
        return (int(parts[1]), int(parts[0]))
    except (ValueError, IndexError):
        return (9999, 99)


def gerar_planilha_p300(pasta_entrada: str, callback_log=None):
    """
    Lê todos os PDFs da pasta, extrai ocorrências da verba P300
    e gera uma planilha .xlsx com uma linha por ocorrência.

    Colunas: Contracheque | Competência | Valor (R$) | Observação

    Retorna (sucesso: bool, mensagem: str).
    """
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    def log(msg):
        if callback_log:
            callback_log(msg)

    if not os.path.isdir(pasta_entrada):
        return False, f"Pasta '{pasta_entrada}' não encontrada."

    arquivos_pdf = sorted([
        os.path.join(pasta_entrada, f)
        for f in os.listdir(pasta_entrada)
        if f.lower().endswith('.pdf')
    ])

    if not arquivos_pdf:
        return False, "Nenhum arquivo PDF encontrado na pasta selecionada."

    log(f"Encontrados {len(arquivos_pdf)} arquivo(s) PDF.\n")

    todas_ocorrencias: list[dict] = []
    arquivos_sem_p300 = 0

    for caminho in arquivos_pdf:
        nome = os.path.basename(caminho)
        ocorrencias = extrair_p300_pdf(caminho)
        if ocorrencias:
            todas_ocorrencias.extend(ocorrencias)
            for oc in ocorrencias:
                sinal = " (Estorno)" if oc['valor'] < 0 else ""
                log(
                    f"  P300  {nome}  ->  "
                    f"Folha {oc['contracheque']}  "
                    f"Comp {oc['competencia']}  "
                    f"R$ {oc['valor']:,.2f}{sinal}"
                )
        else:
            arquivos_sem_p300 += 1
            log(f"  ---   {nome}  ->  Sem verba P300")

    if not todas_ocorrencias:
        return False, "Nenhuma ocorrência de P300 encontrada nos contracheques."

    # Ordenar por contracheque (cronológico) mantendo ordem de aparição dentro
    # da mesma folha
    todas_ocorrencias.sort(key=lambda oc: _chave_ordenacao_mes_ano(oc['contracheque']))

    # Nome do cliente a partir do nome da pasta
    nome_pasta = os.path.basename(pasta_entrada.rstrip(os.sep))
    nome_cliente = re.sub(r'[\\/*?:"<>|]', '_', nome_pasta)

    # --- Criar planilha ---
    wb = Workbook()
    ws = wb.active
    ws.title = "Ocorrências P300"

    # Estilos
    borda = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="003E63", end_color="003E63", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center")

    data_font = Font(name="Arial", size=10)
    align_center = Alignment(horizontal="center", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    fill_branco = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_azul = PatternFill(start_color="EBF3FB", end_color="EBF3FB", fill_type="solid")
    fill_estorno = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    # Cabeçalho
    colunas = ["Contracheque", "Competência", "Valor (R$)", "Observação"]
    ws.row_dimensions[1].height = 22
    for col_idx, titulo in enumerate(colunas, 1):
        cell = ws.cell(row=1, column=col_idx, value=titulo)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = borda

    # Dados
    for row_idx, oc in enumerate(todas_ocorrencias, 2):
        is_estorno = oc['valor'] < 0
        observacao = "Estorno" if is_estorno else ""

        # Determinar preenchimento de fundo
        if is_estorno:
            fill = fill_estorno
        elif (row_idx % 2) == 0:
            fill = fill_branco
        else:
            fill = fill_azul

        # Coluna A: Contracheque (MM/AAAA)
        cell = ws.cell(row=row_idx, column=1, value=oc['contracheque'])
        cell.font = data_font
        cell.alignment = align_center
        cell.border = borda
        cell.fill = fill

        # Coluna B: Competência (MM/AAAA)
        cell = ws.cell(row=row_idx, column=2, value=oc['competencia'])
        cell.font = data_font
        cell.alignment = align_center
        cell.border = borda
        cell.fill = fill

        # Coluna C: Valor (R$)
        cell = ws.cell(row=row_idx, column=3, value=oc['valor'])
        cell.number_format = '#,##0.00'
        cell.font = data_font
        cell.alignment = align_right
        cell.border = borda
        cell.fill = fill

        # Coluna D: Observação
        cell = ws.cell(row=row_idx, column=4, value=observacao)
        cell.font = data_font
        cell.alignment = align_left
        cell.border = borda
        cell.fill = fill

    # Larguras de coluna
    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 20

    # Cabeçalho fixo
    ws.freeze_panes = "A2"

    # ------------------------------------------------------------------
    # Aba 2: Consolidado por Competência
    # ------------------------------------------------------------------
    from collections import defaultdict

    # Agrupar por competência a partir dos dados já extraídos
    agrupado: dict[str, list[float]] = defaultdict(list)
    for oc in todas_ocorrencias:
        agrupado[oc['competencia']].append(oc['valor'])

    # Ordenar competências cronologicamente (converter temporariamente para
    # comparação, mas manter como string para exibição)
    competencias_ordenadas = sorted(
        agrupado.keys(),
        key=_chave_ordenacao_mes_ano,
    )

    ws2 = wb.create_sheet(title="Consolidado por Competência")

    # Estilos específicos da aba 2
    fill_negativo = PatternFill(
        start_color="FFE0E0", end_color="FFE0E0", fill_type="solid",
    )

    # Cabeçalho
    colunas2 = ["Competência", "Qtd. Lançamentos", "Valor Resultante (R$)"]
    ws2.row_dimensions[1].height = 22
    for col_idx, titulo in enumerate(colunas2, 1):
        cell = ws2.cell(row=1, column=col_idx, value=titulo)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = borda

    # Dados
    for row_idx, comp in enumerate(competencias_ordenadas, 2):
        lancamentos = agrupado[comp]
        soma = round(sum(lancamentos), 2)
        qtd = len(lancamentos)

        # Fundo: negativo → vermelho claro, senão alternância branco/azul
        if soma < 0:
            fill = fill_negativo
        elif (row_idx % 2) == 0:
            fill = fill_branco
        else:
            fill = fill_azul

        # Coluna A: Competência
        cell = ws2.cell(row=row_idx, column=1, value=comp)
        cell.font = data_font
        cell.alignment = align_center
        cell.border = borda
        cell.fill = fill

        # Coluna B: Qtd. Lançamentos
        cell = ws2.cell(row=row_idx, column=2, value=qtd)
        cell.font = data_font
        cell.alignment = align_center
        cell.border = borda
        cell.fill = fill

        # Coluna C: Valor Resultante (R$)
        cell = ws2.cell(row=row_idx, column=3, value=soma)
        cell.number_format = '#,##0.00'
        cell.font = data_font
        cell.alignment = align_right
        cell.border = borda
        cell.fill = fill

    # Linha de totais
    total_row = len(competencias_ordenadas) + 2
    total_competencias = len(competencias_ordenadas)
    soma_geral = round(sum(
        sum(agrupado[c]) for c in competencias_ordenadas
    ), 2)

    total_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)

    cell = ws2.cell(row=total_row, column=1, value=f"Total ({total_competencias})")
    cell.font = total_font
    cell.fill = header_fill
    cell.alignment = align_center
    cell.border = borda

    cell = ws2.cell(
        row=total_row, column=2,
        value=sum(len(agrupado[c]) for c in competencias_ordenadas),
    )
    cell.font = total_font
    cell.fill = header_fill
    cell.alignment = align_center
    cell.border = borda

    cell = ws2.cell(row=total_row, column=3, value=soma_geral)
    cell.number_format = '#,##0.00'
    cell.font = total_font
    cell.fill = header_fill
    cell.alignment = align_right
    cell.border = borda

    # Larguras
    ws2.column_dimensions['A'].width = 18
    ws2.column_dimensions['B'].width = 22
    ws2.column_dimensions['C'].width = 24

    # Cabeçalho fixo
    ws2.freeze_panes = "A2"

    nome_arquivo = f"P300_{nome_cliente}.xlsx"
    caminho_saida = os.path.join(pasta_entrada, nome_arquivo)
    wb.save(caminho_saida)

    total_estornos = sum(1 for oc in todas_ocorrencias if oc['valor'] < 0)
    resumo = (
        f"\nPlanilha P300 gerada com sucesso!\n"
        f"  Arquivo: {nome_arquivo}\n"
        f"  Aba 1 - Ocorrências P300: {len(todas_ocorrencias)} linhas\n"
        f"  Aba 2 - Consolidado por Competência: "
        f"{total_competencias} competências\n"
        f"  Estornos: {total_estornos}\n"
        f"  Arquivos sem P300: {arquivos_sem_p300}"
    )

    log(resumo)
    return True, caminho_saida


# ---------------------------------------------------------------------------
# Interface gráfica
# ---------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Organizador de Contracheques")
        self.geometry("750x550")
        self.resizable(True, True)
        self.configure(bg=COR_FUNDO)
        self.minsize(600, 450)

        self._criar_estilos()
        self._criar_widgets()

    def _criar_estilos(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        # Frame
        style.configure("TFrame", background=COR_FUNDO)

        # Labels
        style.configure(
            "TLabel",
            background=COR_FUNDO,
            font=("Segoe UI", 10),
        )
        style.configure(
            "Header.TLabel",
            background=COR_PRIMARIA,
            foreground=COR_TEXTO_CLARO,
            font=("Segoe UI", 14, "bold"),
            padding=(15, 12),
        )
        style.configure(
            "Section.TLabel",
            background=COR_FUNDO,
            foreground=COR_PRIMARIA,
            font=("Segoe UI", 10, "bold"),
        )

        # Entry
        style.configure(
            "TEntry",
            fieldbackground=COR_FUNDO_LOG,
            padding=5,
        )

        # Botão padrão (selecionar pasta)
        style.configure(
            "TButton",
            font=("Segoe UI", 10),
            padding=(12, 6),
        )

        # Botão primário (ações principais)
        style.configure(
            "Primary.TButton",
            font=("Segoe UI", 11, "bold"),
            padding=(16, 10),
            background=COR_PRIMARIA,
            foreground=COR_TEXTO_CLARO,
        )
        style.map(
            "Primary.TButton",
            background=[
                ("active", COR_PRIMARIA_HOVER),
                ("disabled", "#a0aab4"),
            ],
            foreground=[
                ("disabled", "#e0e0e0"),
            ],
        )

        # Botão secundário
        style.configure(
            "Secondary.TButton",
            font=("Segoe UI", 11, "bold"),
            padding=(16, 10),
            background="#1a7a3a",
            foreground=COR_TEXTO_CLARO,
        )
        style.map(
            "Secondary.TButton",
            background=[
                ("active", "#1e8c42"),
                ("disabled", "#a0aab4"),
            ],
            foreground=[
                ("disabled", "#e0e0e0"),
            ],
        )

    def _criar_widgets(self):
        # --- Cabeçalho ---
        header = ttk.Label(
            self,
            text="Organizador de Contracheques",
            style="Header.TLabel",
        )
        header.pack(fill=tk.X)

        # --- Separador visual ---
        sep = tk.Frame(self, height=3, bg=COR_PRIMARIA)
        sep.pack(fill=tk.X)

        # --- Seção da pasta ---
        frame_pasta = ttk.Frame(self, padding=(15, 15, 15, 5))
        frame_pasta.pack(fill=tk.X)

        ttk.Label(
            frame_pasta,
            text="Pasta com os contracheques:",
            style="Section.TLabel",
        ).pack(anchor=tk.W)

        frame_input = ttk.Frame(frame_pasta)
        frame_input.pack(fill=tk.X, pady=(8, 0))

        self.var_pasta = tk.StringVar()
        self.entry_pasta = ttk.Entry(frame_input, textvariable=self.var_pasta)
        self.entry_pasta.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.btn_selecionar = ttk.Button(
            frame_input, text="Selecionar pasta...", command=self.selecionar_pasta
        )
        self.btn_selecionar.pack(side=tk.LEFT, padx=(8, 0))

        # --- Botões de ação ---
        frame_btn = ttk.Frame(self, padding=(15, 10))
        frame_btn.pack(fill=tk.X)

        self.btn_compilado = ttk.Button(
            frame_btn,
            text="Gerar Compilado",
            style="Primary.TButton",
            command=self.iniciar_compilado,
        )
        self.btn_compilado.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5))

        self.btn_p300 = ttk.Button(
            frame_btn,
            text="Gerar Planilha P300",
            style="Secondary.TButton",
            command=self.iniciar_p300,
        )
        self.btn_p300.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(5, 0))

        # --- Área de log ---
        frame_log = ttk.Frame(self, padding=(15, 5, 15, 15))
        frame_log.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame_log, text="Progresso:", style="Section.TLabel").pack(anchor=tk.W)

        log_container = tk.Frame(frame_log, bg=COR_BORDA, bd=1, relief=tk.SOLID)
        log_container.pack(fill=tk.BOTH, expand=True, pady=(8, 0))

        self.texto_log = tk.Text(
            log_container,
            height=15,
            state=tk.DISABLED,
            wrap=tk.WORD,
            font=("Consolas", 9),
            bg=COR_FUNDO_LOG,
            fg="#1a1a2e",
            padx=10,
            pady=8,
            relief=tk.FLAT,
            borderwidth=0,
        )
        scrollbar = ttk.Scrollbar(log_container, command=self.texto_log.yview)
        self.texto_log.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.texto_log.pack(fill=tk.BOTH, expand=True)

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta com os contracheques")
        if pasta:
            self.var_pasta.set(pasta)

    def adicionar_log(self, msg):
        self.texto_log.configure(state=tk.NORMAL)
        self.texto_log.insert(tk.END, msg + "\n")
        self.texto_log.see(tk.END)
        self.texto_log.configure(state=tk.DISABLED)

    def _validar_pasta(self) -> str | None:
        pasta = self.var_pasta.get().strip()
        if not pasta:
            messagebox.showwarning("Atenção", "Selecione uma pasta primeiro.")
            return None
        if not os.path.isdir(pasta):
            messagebox.showerror("Erro", f"Pasta não encontrada:\n{pasta}")
            return None
        return pasta

    def _limpar_log(self):
        self.texto_log.configure(state=tk.NORMAL)
        self.texto_log.delete("1.0", tk.END)
        self.texto_log.configure(state=tk.DISABLED)

    def _desabilitar_botoes(self):
        self.btn_compilado.configure(state=tk.DISABLED)
        self.btn_p300.configure(state=tk.DISABLED)
        self.btn_selecionar.configure(state=tk.DISABLED)

    def _habilitar_botoes(self):
        self.btn_compilado.configure(state=tk.NORMAL)
        self.btn_p300.configure(state=tk.NORMAL)
        self.btn_selecionar.configure(state=tk.NORMAL)

    # --- Gerar Compilado ---
    def iniciar_compilado(self):
        pasta = self._validar_pasta()
        if not pasta:
            return
        self._limpar_log()
        self._desabilitar_botoes()
        thread = threading.Thread(target=self._processar_compilado, args=(pasta,), daemon=True)
        thread.start()

    def _processar_compilado(self, pasta):
        def log_thread_safe(msg):
            self.after(0, self.adicionar_log, msg)

        sucesso, resultado = organizar_contracheques(pasta, callback_log=log_thread_safe)

        def finalizar():
            self._habilitar_botoes()
            if sucesso:
                messagebox.showinfo("Sucesso", f"Compilado salvo em:\n{resultado}")
            else:
                messagebox.showerror("Erro", resultado)

        self.after(0, finalizar)

    # --- Gerar Planilha P300 ---
    def iniciar_p300(self):
        pasta = self._validar_pasta()
        if not pasta:
            return
        self._limpar_log()
        self._desabilitar_botoes()
        thread = threading.Thread(target=self._processar_p300, args=(pasta,), daemon=True)
        thread.start()

    def _processar_p300(self, pasta):
        def log_thread_safe(msg):
            self.after(0, self.adicionar_log, msg)

        sucesso, resultado = gerar_planilha_p300(pasta, callback_log=log_thread_safe)

        def finalizar():
            self._habilitar_botoes()
            if sucesso:
                messagebox.showinfo("Sucesso", f"Planilha salva em:\n{resultado}")
            else:
                messagebox.showerror("Erro", resultado)

        self.after(0, finalizar)


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
