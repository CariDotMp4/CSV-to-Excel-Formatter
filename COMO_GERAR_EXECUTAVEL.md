# Como gerar o executável (.exe no Windows / binário no Mac/Linux)

## 1. Instale as dependências

```bash
pip install pandas openpyxl pyinstaller
```

## 2. Gere o executável

### Windows (gera csv_to_excel.exe — sem janela de terminal):
```bash
pyinstaller --onefile --windowed --name "CSV_Excel_Formatter" csv_to_excel.py
```

### Mac / Linux (gera um binário):
```bash
pyinstaller --onefile --name "csv_to_excel" csv_to_excel.py
```

O executável estará na pasta `dist/`.

---

## Modos de uso do script

### Interface gráfica (padrão — só dar duplo clique no .exe):
```bash
python csv_to_excel.py
```

### Linha de comando:
```bash
python csv_to_excel.py --cli \
  --csv "C:\SharePoint\Pasta\dados.csv" \
  --saida "C:\Relatorios\resultado.xlsx" \
  --titulo "Meu Relatório" \
  --sep ";"
```

| Parâmetro | Descrição                        | Padrão  |
|-----------|----------------------------------|---------|
| `--csv`   | Caminho do arquivo CSV           | —       |
| `--saida` | Caminho do Excel de saída        | —       |
| `--titulo`| Título da planilha               | "Relatório de Dados" |
| `--sep`   | Separador: `,`  `;`  `\t`        | `,`     |
| `--enc`   | Encoding (utf-8, latin-1…)       | utf-8   |
