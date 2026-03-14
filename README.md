# Organizador de Contracheques

Consolida múltiplos PDFs de contracheques em um único arquivo ordenado cronologicamente (ordem crescente por mês/ano).

Basta abrir o programa, selecionar a pasta com os PDFs e clicar em "Gerar compilado". O arquivo de saída é salvo na mesma pasta.

## Como gerar o .exe

1. Instale o Python (marque "Add Python to PATH" na instalação)
2. Abra o Prompt de Comando na pasta do projeto e execute:

```bash
pip install -r requirements.txt
pyinstaller --onefile --windowed organizar_contracheques.py
```

3. O `.exe` será gerado em `dist/organizar_contracheques.exe`
4. Copie esse arquivo para onde quiser e use com dois cliques

## Como funciona

1. Abre uma janela para você selecionar a pasta com os contracheques
2. Extrai a data (mês/ano) de cada PDF automaticamente
3. Ordena em ordem cronológica crescente
4. Gera um único PDF compilado na mesma pasta de entrada

### Formatos de data suportados

- **Numérico**: `12/2023`, `06/2020`, `02.2026`
- **Competência/Referência**: `Competência: 12/2023`
- **Por extenso**: `Janeiro de 2023`, `FEVEREIRO/2024`

### Alertas

- Arquivos com data não identificada são adicionados ao final do compilado
- Duplicatas (mesmo mês/ano) são sinalizadas na tela
