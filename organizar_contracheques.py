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
    pip install pdfplumber pypdf

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


# Padrões comuns de mês/ano encontrados em contracheques
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


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Organizador de Contracheques")
        self.geometry("700x500")
        self.resizable(True, True)

        # --- Seção da pasta ---
        frame_pasta = ttk.Frame(self, padding=10)
        frame_pasta.pack(fill=tk.X)

        ttk.Label(frame_pasta, text="Pasta com os contracheques:").pack(anchor=tk.W)

        frame_input = ttk.Frame(frame_pasta)
        frame_input.pack(fill=tk.X, pady=(5, 0))

        self.var_pasta = tk.StringVar()
        self.entry_pasta = ttk.Entry(frame_input, textvariable=self.var_pasta)
        self.entry_pasta.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.btn_selecionar = ttk.Button(
            frame_input, text="Selecionar pasta...", command=self.selecionar_pasta
        )
        self.btn_selecionar.pack(side=tk.LEFT, padx=(5, 0))

        # --- Botão processar ---
        frame_btn = ttk.Frame(self, padding=(10, 5))
        frame_btn.pack(fill=tk.X)

        self.btn_processar = ttk.Button(
            frame_btn, text="Gerar compilado", command=self.iniciar_processamento
        )
        self.btn_processar.pack(fill=tk.X)

        # --- Área de log ---
        frame_log = ttk.Frame(self, padding=10)
        frame_log.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame_log, text="Progresso:").pack(anchor=tk.W)

        self.texto_log = tk.Text(frame_log, height=15, state=tk.DISABLED, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(frame_log, command=self.texto_log.yview)
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

    def iniciar_processamento(self):
        pasta = self.var_pasta.get().strip()
        if not pasta:
            messagebox.showwarning("Atenção", "Selecione uma pasta primeiro.")
            return

        if not os.path.isdir(pasta):
            messagebox.showerror("Erro", f"Pasta não encontrada:\n{pasta}")
            return

        # Limpar log
        self.texto_log.configure(state=tk.NORMAL)
        self.texto_log.delete("1.0", tk.END)
        self.texto_log.configure(state=tk.DISABLED)

        # Desabilitar botões durante processamento
        self.btn_processar.configure(state=tk.DISABLED)
        self.btn_selecionar.configure(state=tk.DISABLED)

        # Processar em thread separada para não travar a janela
        thread = threading.Thread(target=self.processar, args=(pasta,), daemon=True)
        thread.start()

    def processar(self, pasta):
        def log_thread_safe(msg):
            self.after(0, self.adicionar_log, msg)

        sucesso, resultado = organizar_contracheques(pasta, callback_log=log_thread_safe)

        def finalizar():
            self.btn_processar.configure(state=tk.NORMAL)
            self.btn_selecionar.configure(state=tk.NORMAL)
            if sucesso:
                messagebox.showinfo(
                    "Sucesso",
                    f"Compilado salvo em:\n{resultado}"
                )
            else:
                messagebox.showerror("Erro", resultado)

        self.after(0, finalizar)


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
