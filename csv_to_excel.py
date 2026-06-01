"""
csv_to_excel.py
---------------
Lê um CSV (de pasta local do SharePoint ou qualquer pasta) e gera
uma planilha Excel com formatação roxa profissional.

Modos de uso:
  • Interface gráfica (padrão):   python csv_to_excel.py
  • Linha de comando:             python csv_to_excel.py --cli --csv ... --saida ... --titulo ...
  • Gerar executável:             pyinstaller --onefile --windowed csv_to_excel.py
"""

import argparse
import os
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ══════════════════════════════════════════════════════════════════════════════
#  ESTILOS DA PLANILHA
# ══════════════════════════════════════════════════════════════════════════════

class Estilos:
    ROXO            = "7030A0"   # cabeçalho e rodapé
    ROXO_CLARO      = "E2D0F0"   # linha par (zebra)
    BRANCO          = "FFFFFF"
    PRETO           = "000000"
    BORDA           = "7030A0"   # bordas roxas finas

    FONTE           = "Arial"
    T_TITULO        = 16
    T_SUBTITULO     = 10
    T_CABECALHO     = 11
    T_DADOS         = 10
    T_RODAPE        = 10

    L_TITULO        = 1
    L_SUBTITULO     = 2
    L_BRANCO        = 3
    L_CABECALHO     = 4


# ══════════════════════════════════════════════════════════════════════════════
#  HELPERS DE ESTILO OPENPYXL
# ══════════════════════════════════════════════════════════════════════════════

def _borda(cor: str = Estilos.BORDA) -> Border:
    s = Side(style="thin", color=cor)
    return Border(left=s, right=s, top=s, bottom=s)

def _fill(cor: str) -> PatternFill:
    return PatternFill("solid", fgColor=cor)

def _font(cor: str, size: int, bold: bool = False) -> Font:
    return Font(name=Estilos.FONTE, size=size, bold=bold, color=cor)

def _align(h: str = "left", v: str = "center", wrap: bool = False) -> Alignment:
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


# ══════════════════════════════════════════════════════════════════════════════
#  LEITURA DO CSV
# ══════════════════════════════════════════════════════════════════════════════

def ler_csv(caminho: Path, sep: str = ",", enc: str = "utf-8") -> pd.DataFrame:
    try:
        df = pd.read_csv(caminho, sep=sep, encoding=enc, skipinitialspace=True)
    except UnicodeDecodeError:
        df = pd.read_csv(caminho, sep=sep, encoding="latin-1", skipinitialspace=True)
    df.dropna(how="all", axis=1, inplace=True)
    df.dropna(how="all", axis=0, inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df


# ══════════════════════════════════════════════════════════════════════════════
#  GERAÇÃO DO EXCEL
# ══════════════════════════════════════════════════════════════════════════════

def _col_width(df: pd.DataFrame, col: str) -> float:
    lens = [len(str(col))] + [len(str(v)) for v in df[col] if pd.notna(v)]
    return min(max(max(lens) + 2, 10), 60)


def gerar_excel(df: pd.DataFrame, caminho_saida: str,
                titulo: str = "Relatório de Dados") -> str:
    wb = Workbook()
    ws = wb.active
    ws.title = "Dados"

    n_cols   = len(df.columns)
    n_linhas = len(df)
    col_fim  = get_column_letter(n_cols)

    # ── Título mesclado ──────────────────────────────────────────────────────
    ws.merge_cells(f"A{Estilos.L_TITULO}:{col_fim}{Estilos.L_TITULO}")
    ct = ws[f"A{Estilos.L_TITULO}"]
    ct.value     = titulo
    ct.font      = _font(Estilos.ROXO, Estilos.T_TITULO, bold=True)
    ct.alignment = _align("center")
    ws.row_dimensions[Estilos.L_TITULO].height = 34

    # ── Subtítulo: contagem de registros ─────────────────────────────────────
    ws.merge_cells(f"A{Estilos.L_SUBTITULO}:{col_fim}{Estilos.L_SUBTITULO}")
    cs = ws[f"A{Estilos.L_SUBTITULO}"]
    cs.value     = f"{n_linhas} registro(s)  ·  {n_cols} coluna(s)"
    cs.font      = _font(Estilos.ROXO, Estilos.T_SUBTITULO)
    cs.alignment = _align("center")
    ws.row_dimensions[Estilos.L_SUBTITULO].height = 18

    # ── Linha de respiro ──────────────────────────────────────────────────────
    ws.row_dimensions[Estilos.L_BRANCO].height = 8

    # ── Cabeçalho da tabela ──────────────────────────────────────────────────
    for c, nome in enumerate(df.columns, 1):
        cel = ws.cell(row=Estilos.L_CABECALHO, column=c, value=str(nome).strip())
        cel.font      = _font(Estilos.BRANCO, Estilos.T_CABECALHO, bold=True)
        cel.fill      = _fill(Estilos.ROXO)
        cel.alignment = _align("center", wrap=True)
        cel.border    = _borda()
    ws.row_dimensions[Estilos.L_CABECALHO].height = 30

    # ── Dados com zebra ──────────────────────────────────────────────────────
    l0 = Estilos.L_CABECALHO + 1
    for r, row in enumerate(df.itertuples(index=False)):
        linha = l0 + r
        fundo = Estilos.ROXO_CLARO if r % 2 == 0 else Estilos.BRANCO
        for c, val in enumerate(row, 1):
            v   = "" if pd.isna(val) else val
            cel = ws.cell(row=linha, column=c, value=v)
            cel.font   = _font(Estilos.PRETO, Estilos.T_DADOS)
            cel.fill   = _fill(fundo)
            cel.border = _borda()
            cel.alignment = (
                _align("right") if isinstance(val, (int, float)) and pd.notna(val)
                else _align("left")
            )
        ws.row_dimensions[linha].height = 18

    # ── Rodapé de totais ─────────────────────────────────────────────────────
    lr = l0 + n_linhas
    ws.row_dimensions[lr].height = 22
    for c, nome in enumerate(df.columns, 1):
        cel        = ws.cell(row=lr, column=c)
        cel.fill   = _fill(Estilos.ROXO)
        cel.font   = _font(Estilos.BRANCO, Estilos.T_RODAPE, bold=True)
        cel.border = _borda()
        cel.alignment = _align("center")
        if pd.api.types.is_numeric_dtype(df[nome]):
            cl        = get_column_letter(c)
            cel.value = f"=SUM({cl}{l0}:{cl}{lr-1})"
        elif c == 1:
            cel.value = f"Total: {n_linhas} registros"

    # ── Larguras automáticas ─────────────────────────────────────────────────
    for c, nome in enumerate(df.columns, 1):
        ws.column_dimensions[get_column_letter(c)].width = _col_width(df, nome)

    # ── Congelar cabeçalho + filtros ─────────────────────────────────────────
    ws.freeze_panes = f"A{l0}"
    ws.auto_filter.ref = f"A{Estilos.L_CABECALHO}:{col_fim}{lr-1}"

    # ── Impressão em paisagem ─────────────────────────────────────────────────
    ws.page_setup.orientation  = "landscape"
    ws.page_setup.fitToPage    = True
    ws.page_setup.fitToWidth   = 1
    ws.page_setup.fitToHeight  = 0
    ws.print_title_rows = f"{Estilos.L_CABECALHO}:{Estilos.L_CABECALHO}"

    # ── Aba Info ──────────────────────────────────────────────────────────────
    wi = wb.create_sheet("Info")
    for i, (k, v) in enumerate([
        ("Gerado por", "csv_to_excel"),
        ("Registros",  n_linhas),
        ("Colunas",    n_cols),
        ("Lista de colunas", ", ".join(df.columns.tolist())),
    ], 1):
        wi.cell(i, 1, k).font = Font(bold=True)
        wi.cell(i, 2, v)
    wi.column_dimensions["A"].width = 20
    wi.column_dimensions["B"].width = 70

    wb.save(caminho_saida)
    return caminho_saida


# ══════════════════════════════════════════════════════════════════════════════
#  INTERFACE GRÁFICA (TKINTER)
# ══════════════════════════════════════════════════════════════════════════════

ROXO_UI      = "#7030A0"
ROXO_HOVER   = "#5a2080"
ROXO_CLARO   = "#f3e8fc"
ROXO_BG      = "#faf5ff"
TEXTO        = "#1a1a2e"
CINZA_BORDA  = "#d8b4fe"
BRANCO       = "#ffffff"
VERDE_OK     = "#16a34a"
VERMELHO_ERR = "#dc2626"

PH_CSV    = "Caminho completo do arquivo .csv"
PH_SAIDA  = "Caminho onde o Excel será salvo"
PH_TITULO = "Ex: Relatório de Vendas 2025"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CSV → Excel Formatter")
        self.resizable(False, False)
        self.configure(bg=ROXO_BG)
        self._center(680, 520)
        self._build()

    def _center(self, w, h):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    # ── Construção da UI ─────────────────────────────────────────────────────

    def _build(self):
        # ── Cabeçalho ────────────────────────────────────────────────────────
        header = tk.Frame(self, bg=ROXO_UI, height=72)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="CSV → Excel Formatter",
                 bg=ROXO_UI, fg=BRANCO,
                 font=("Arial", 18, "bold")).pack(side="left", padx=24, pady=16)
        tk.Label(header, text="SharePoint / Local",
                 bg=ROXO_UI, fg="#d8b4fe",
                 font=("Arial", 10)).pack(side="right", padx=24)

        # ── Corpo ─────────────────────────────────────────────────────────────
        body = tk.Frame(self, bg=ROXO_BG, padx=32, pady=20)
        body.pack(fill="both", expand=True)

        # Variáveis dos campos
        self.csv_path  = tk.StringVar()
        self.saida_path = tk.StringVar()
        self.titulo_var = tk.StringVar()
        self.sep_var    = tk.StringVar(value=",")

        # Campo 1 — CSV de entrada
        self._bloco(body, "📄  Arquivo CSV de entrada",
                    self.csv_path, PH_CSV,
                    btn_texto="Procurar…", btn_cmd=self._browse_csv)

        # Campo 2 — Saída xlsx
        self._bloco(body, "💾  Arquivo de saída (.xlsx)",
                    self.saida_path, PH_SAIDA,
                    btn_texto="Salvar em…", btn_cmd=self._browse_saida)

        # Campo 3 — Título
        self._bloco(body, "🏷️  Título da planilha",
                    self.titulo_var, PH_TITULO)

        # Campo 4 — Separador (linha curta, sem botão)
        frm_sep = tk.Frame(body, bg=ROXO_BG)
        frm_sep.pack(fill="x", pady=(10, 0))
        tk.Label(frm_sep, text="🔤  Separador do CSV",
                 bg=ROXO_BG, fg=TEXTO,
                 font=("Arial", 9, "bold")).pack(anchor="w")
        tk.Entry(frm_sep, textvariable=self.sep_var,
                 width=6, font=("Arial", 10),
                 bg=BRANCO, fg=TEXTO, relief="flat",
                 highlightthickness=1,
                 highlightbackground=CINZA_BORDA,
                 highlightcolor=ROXO_UI,
                 insertbackground=ROXO_UI
                 ).pack(anchor="w", ipady=5, pady=(4, 0))
        tk.Label(frm_sep, text="  Use  ,  para vírgula  |  ;  para ponto-e-vírgula  |  \\t  para tabulação",
                 bg=ROXO_BG, fg="#9ca3af", font=("Arial", 8)).pack(anchor="w")

        # ── Barra de status ───────────────────────────────────────────────────
        self.status_var = tk.StringVar(value="Preencha os campos e clique em Gerar.")
        self.status_lbl = tk.Label(body, textvariable=self.status_var,
                                   bg=ROXO_BG, fg=TEXTO,
                                   font=("Arial", 9), anchor="w", wraplength=600)
        self.status_lbl.pack(fill="x", pady=(18, 4))

        # ── Barra de progresso ────────────────────────────────────────────────
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("Purple.Horizontal.TProgressbar",
                        troughcolor=ROXO_CLARO,
                        background=ROXO_UI,
                        thickness=6)
        self.progress = ttk.Progressbar(body, mode="indeterminate",
                                        style="Purple.Horizontal.TProgressbar")
        self.progress.pack(fill="x")

        # ── Botão gerar ───────────────────────────────────────────────────────
        self.btn = tk.Button(body, text="✨  Gerar Planilha",
                             bg=ROXO_UI, fg=BRANCO,
                             font=("Arial", 12, "bold"),
                             relief="flat", cursor="hand2",
                             activebackground=ROXO_HOVER, activeforeground=BRANCO,
                             padx=28, pady=10,
                             command=self._iniciar)
        self.btn.pack(pady=(18, 0))
        self.btn.bind("<Enter>", lambda e: self.btn.config(bg=ROXO_HOVER))
        self.btn.bind("<Leave>", lambda e: self.btn.config(bg=ROXO_UI))

    def _bloco(self, parent, label: str, var: tk.StringVar,
               placeholder: str, btn_texto: str = None, btn_cmd=None):
        """Cria um bloco label + (entry + botão opcional) usando pack."""
        frm = tk.Frame(parent, bg=ROXO_BG)
        frm.pack(fill="x", pady=(10, 0))

        tk.Label(frm, text=label, bg=ROXO_BG, fg=TEXTO,
                 font=("Arial", 9, "bold")).pack(anchor="w")

        row = tk.Frame(frm, bg=ROXO_BG)
        row.pack(fill="x", pady=(4, 0))

        entry = tk.Entry(row, textvariable=var,
                         font=("Arial", 10),
                         bg=BRANCO, fg="#9ca3af",
                         relief="flat",
                         highlightthickness=1,
                         highlightbackground=CINZA_BORDA,
                         highlightcolor=ROXO_UI,
                         insertbackground=ROXO_UI)
        entry.pack(side="left", fill="x", expand=True, ipady=6)

        # Placeholder
        entry.insert(0, placeholder)
        entry.bind("<FocusIn>",  lambda e, en=entry, ph=placeholder: self._ph_in(en, ph))
        entry.bind("<FocusOut>", lambda e, en=entry, ph=placeholder: self._ph_out(en, ph))

        if btn_texto and btn_cmd:
            tk.Button(row, text=btn_texto,
                      bg=ROXO_CLARO, fg=ROXO_UI,
                      font=("Arial", 9, "bold"), relief="flat", cursor="hand2",
                      activebackground=CINZA_BORDA, activeforeground=ROXO_UI,
                      padx=10,
                      command=btn_cmd
                      ).pack(side="left", padx=(8, 0), ipady=6)

    # ── Placeholder ──────────────────────────────────────────────────────────

    def _ph_in(self, entry, ph):
        if entry.get() == ph:
            entry.delete(0, "end")
            entry.config(fg=TEXTO)

    def _ph_out(self, entry, ph):
        if not entry.get():
            entry.insert(0, ph)
            entry.config(fg="#9ca3af")

    # ── Diálogos de arquivo ──────────────────────────────────────────────────

    def _browse_csv(self):
        p = filedialog.askopenfilename(
            parent=self,
            title="Selecionar arquivo CSV",
            filetypes=[("CSV", "*.csv"), ("Todos os arquivos", "*.*")]
        )
        if not p:
            return
        self.csv_path.set(p)
        # Sugere saída automática
        if self.saida_path.get() in ("", PH_SAIDA):
            self.saida_path.set(str(Path(p).with_suffix(".xlsx")))
        # Sugere título automático
        if self.titulo_var.get() in ("", PH_TITULO):
            self.titulo_var.set(Path(p).stem.replace("_", " ").title())

    def _browse_saida(self):
        # Tenta iniciar no diretório do CSV, se já selecionado
        csv = self.csv_path.get()
        inicio = str(Path(csv).parent) if csv and csv != PH_CSV and Path(csv).exists() else os.path.expanduser("~")

        p = filedialog.asksaveasfilename(
            parent=self,
            title="Salvar planilha Excel como…",
            initialdir=inicio,
            initialfile="relatorio.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel (.xlsx)", "*.xlsx"), ("Todos os arquivos", "*.*")]
        )
        if not p:
            return
        # Garante extensão correta
        if not p.lower().endswith(".xlsx"):
            p += ".xlsx"
        self.saida_path.set(p)

    # ── Geração ──────────────────────────────────────────────────────────────

    def _set_status(self, msg: str, cor: str = TEXTO):
        self.status_var.set(msg)
        self.status_lbl.config(fg=cor)

    def _iniciar(self):
        csv_path = self.csv_path.get().strip()
        saida    = self.saida_path.get().strip()
        titulo   = self.titulo_var.get().strip()
        sep      = self.sep_var.get().strip() or ","

        if not csv_path or csv_path == PH_CSV:
            self._set_status("⚠️  Informe o caminho do CSV.", VERMELHO_ERR); return
        if not Path(csv_path).exists():
            self._set_status(f"⚠️  Arquivo não encontrado: {csv_path}", VERMELHO_ERR); return
        if not saida or saida == PH_SAIDA:
            self._set_status("⚠️  Informe o caminho de saída.", VERMELHO_ERR); return
        if not titulo or titulo == PH_TITULO:
            titulo = Path(csv_path).stem.replace("_", " ").title()

        sep = sep.replace("\\t", "\t")

        self.btn.config(state="disabled")
        self.progress.start(12)
        self._set_status("⏳  Processando…", ROXO_UI)

        def run():
            try:
                df = ler_csv(Path(csv_path), sep=sep)
                gerar_excel(df, saida, titulo)
                self.after(0, self._sucesso, saida, len(df), len(df.columns))
            except Exception as exc:
                self.after(0, self._erro, str(exc))

        threading.Thread(target=run, daemon=True).start()

    def _sucesso(self, saida: str, linhas: int, colunas: int):
        self.progress.stop()
        self.btn.config(state="normal")
        self._set_status(
            f"✅  Planilha gerada!  {linhas} linhas · {colunas} colunas  →  {saida}",
            VERDE_OK
        )
        if messagebox.askyesno("Concluído",
                               f"Planilha gerada com sucesso!\n\n{saida}\n\nDeseja abrir o arquivo?"):
            import subprocess, platform
            if platform.system() == "Windows":
                os.startfile(saida)
            elif platform.system() == "Darwin":
                subprocess.call(["open", saida])
            else:
                subprocess.call(["xdg-open", saida])

    def _erro(self, msg: str):
        self.progress.stop()
        self.btn.config(state="normal")
        self._set_status(f"❌  Erro: {msg}", VERMELHO_ERR)
        messagebox.showerror("Erro", msg)


# ══════════════════════════════════════════════════════════════════════════════
#  MODO CLI
# ══════════════════════════════════════════════════════════════════════════════

def cli():
    parser = argparse.ArgumentParser(
        description="CSV → Excel formatter (modo terminal)"
    )
    parser.add_argument("--csv",    required=True, help="Caminho do CSV de entrada")
    parser.add_argument("--saida",  required=True, help="Caminho do Excel de saída (.xlsx)")
    parser.add_argument("--titulo", default="Relatório de Dados", help="Título da planilha")
    parser.add_argument("--sep",    default=",", help="Separador do CSV (padrão: vírgula)")
    parser.add_argument("--enc",    default="utf-8", help="Encoding do CSV")
    args = parser.parse_args()

    caminho = Path(args.csv)
    if not caminho.exists():
        print(f"❌  Arquivo não encontrado: {caminho}"); sys.exit(1)

    print(f"🔄  Lendo {caminho.name}…")
    df = ler_csv(caminho, sep=args.sep.replace("\\t", "\t"), enc=args.enc)
    print(f"   ✅  {len(df)} linhas × {len(df.columns)} colunas")

    print("📊  Gerando Excel…")
    gerar_excel(df, args.saida, args.titulo)
    print(f"✅  Salvo em: {args.saida}")


# ══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    if "--cli" in sys.argv:
        sys.argv.remove("--cli")
        cli()
    else:
        App().mainloop()
