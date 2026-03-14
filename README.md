# Organizador de Contracheques

Consolida múltiplos PDFs de contracheques em um único arquivo ordenado cronologicamente (ordem crescente por mês/ano).

## Instalação

```bash
pip install -r requirements.txt
```

## Uso

```bash
python organizar_contracheques.py <PASTA_COM_PDFS> [ARQUIVO_SAIDA]
```

### Exemplos

```bash
# Gera o compilado na mesma pasta de entrada
python organizar_contracheques.py ./contracheques

# Especifica o arquivo de saída
python organizar_contracheques.py ./contracheques ./resultado.pdf
```

## Como funciona

1. Lê todos os PDFs da pasta informada
2. Extrai a data (mês/ano) de cada contracheque automaticamente
3. Ordena em ordem cronológica crescente
4. Gera um único PDF compilado

### Formatos suportados

O programa reconhece datas em diversos formatos:
- **Numérico**: `12/2023`, `06/2020`, `02.2026`
- **Competência/Referência**: `Competência: 12/2023`
- **Por extenso**: `Janeiro de 2023`, `FEVEREIRO/2024`

### Alertas

- Arquivos com data não identificada são adicionados ao final do compilado
- Duplicatas (mesmo mês/ano) são sinalizadas no console
