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

# Padrão para capturar a verba P300 e seu valor
PADRAO_P300 = re.compile(
    r'P\s*300\b.*?(\d[\d.,]*)',
    re.IGNORECASE,
)


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


def extrair_p300_contracheque(caminho_pdf: str) -> list[tuple[datetime | None, float]]:
    """
    Extrai todas as ocorrências da verba P300 de um PDF de contracheque.
    Retorna lista de (data, valor).
    """
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            if not pdf.pages:
                return []
            texto = ""
            for page in pdf.pages:
                texto += (page.extract_text() or "") + "\n"
    except Exception:
        return []

    data = extrair_data_contracheque(caminho_pdf)

    resultados = []
    for match in PADRAO_P300.finditer(texto):
        valor_str = match.group(1).replace('.', '').replace(',', '.')
        try:
            valor = float(valor_str)
            resultados.append((data, valor))
        except ValueError:
            continue

    return resultados


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


def gerar_planilha_p300(pasta_entrada: str, callback_log=None):
    """
    Lê todos os PDFs da pasta, extrai ocorrências da verba P300
    e gera uma planilha .xlsx com a evolução mês a mês.

    Retorna (sucesso: bool, mensagem: str).
    """
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

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

    todas_ocorrencias = []
    arquivos_sem_p300 = 0

    for caminho in arquivos_pdf:
        nome = os.path.basename(caminho)
        ocorrencias = extrair_p300_contracheque(caminho)
        if ocorrencias:
            for data, valor in ocorrencias:
                todas_ocorrencias.append((data, valor, nome))
                if data:
                    log(f"  P300  {nome}  ->  {data.strftime('%m/%Y')}  R$ {valor:,.2f}")
                else:
                    log(f"  P300  {nome}  ->  Data não identificada  R$ {valor:,.2f}")
        else:
            arquivos_sem_p300 += 1
            log(f"  ---   {nome}  ->  Sem verba P300")

    if not todas_ocorrencias:
        return False, "Nenhuma ocorrência de P300 encontrada nos contracheques."

    # Separar com e sem data
    com_data = [(d, v, n) for d, v, n in todas_ocorrencias if d is not None]
    sem_data = [(d, v, n) for d, v, n in todas_ocorrencias if d is None]

    # Ordenar cronologicamente
    com_data.sort(key=lambda x: x[0])

    linhas = com_data + sem_data

    # Tentar extrair nome do cliente a partir do nome da pasta
    nome_pasta = os.path.basename(pasta_entrada.rstrip(os.sep))
    nome_cliente = re.sub(r'[\\/*?:"<>|]', '_', nome_pasta)

    # Criar planilha
    wb = Workbook()
    ws = wb.active
    ws.title = "Evolução P300"

    # Estilos
    header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="003E63", end_color="003E63", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style="thin", color="C8D0D8"),
        right=Side(style="thin", color="C8D0D8"),
        top=Side(style="thin", color="C8D0D8"),
        bottom=Side(style="thin", color="C8D0D8"),
    )
    valor_font = Font(name="Calibri", size=11)
    valor_alignment = Alignment(horizontal="center", vertical="center")

    # Cabeçalho
    colunas = ["Mês/Ano", "Competência", "Valor P300"]
    for col_idx, titulo in enumerate(colunas, 1):
        cell = ws.cell(row=1, column=col_idx, value=titulo)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border

    # Dados
    for row_idx, (data, valor, _nome) in enumerate(linhas, 2):
        if data:
            mes_ano = data.strftime("%m/%Y")
            competencia = data.strftime("%B/%Y").capitalize()
        else:
            mes_ano = "N/D"
            competencia = "Não identificada"

        cell_mes = ws.cell(row=row_idx, column=1, value=mes_ano)
        cell_mes.alignment = valor_alignment
        cell_mes.border = thin_border
        cell_mes.font = valor_font

        cell_comp = ws.cell(row=row_idx, column=2, value=competencia)
        cell_comp.alignment = valor_alignment
        cell_comp.border = thin_border
        cell_comp.font = valor_font

        cell_valor = ws.cell(row=row_idx, column=3, value=valor)
        cell_valor.number_format = '#,##0.00'
        cell_valor.alignment = valor_alignment
        cell_valor.border = thin_border
        cell_valor.font = valor_font

    # Ajustar largura das colunas
    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['C'].width = 18

    nome_arquivo = f"evolucao_P300_{nome_cliente}.xlsx"
    caminho_saida = os.path.join(pasta_entrada, nome_arquivo)
    wb.save(caminho_saida)

    resumo = (
        f"\nPlanilha P300 gerada com sucesso!\n"
        f"  Arquivo: {nome_arquivo}\n"
        f"  Ocorrências de P300: {len(todas_ocorrencias)}\n"
        f"  Arquivos sem P300: {arquivos_sem_p300}"
    )
    if com_data:
        resumo += (
            f"\n  Período: {com_data[0][0].strftime('%m/%Y')} "
            f"a {com_data[-1][0].strftime('%m/%Y')}"
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
