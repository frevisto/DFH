# Última fase do processamento das cotações
# input: Mescla com valores cotados
# output: Cotações separadas, ver outdir/

import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# ===============================
# Processamento principal
# ===============================
def gerar_planilhas_por_cotacao(arquivo_entrada, outdir_base):
    # Nome do arquivo de entrada (sem extensão)
    nome_base = os.path.splitext(os.path.basename(arquivo_entrada))[0]

    # Diretório final: ./outdir/<nome_arquivo_entrada>/
    outdir = os.path.join(outdir_base, nome_base)
    os.makedirs(outdir, exist_ok=True)

    # Ler planilha
    df = pd.read_excel(arquivo_entrada)
    df.columns = df.columns.str.strip()

    # Localizar coluna "Cotação"
    col_lookup = {c.upper(): c for c in df.columns}
    if "COTAÇÃO" not in col_lookup:
        raise ValueError("Coluna 'Cotação' não encontrada.")

    cot_col = col_lookup["COTAÇÃO"]

    # Remover coluna Cotação do conteúdo final
    df_sem_cotacao = df.drop(columns=[cot_col])

    # Agrupamento O(n)
    grupos = df_sem_cotacao.groupby(df[cot_col])

    total = 0

    for cotacao, grupo in grupos:
        if pd.isna(cotacao) or str(cotacao).strip() == "":
            continue

        nome_arquivo = f"{str(cotacao).strip()}.xlsx"
        caminho_saida = os.path.join(outdir, nome_arquivo)

        grupo.to_excel(caminho_saida, index=False)
        total += 1

        print(f"✅ Gerado: {os.path.join(nome_base, nome_arquivo)} ({len(grupo)} linhas)")

    messagebox.showinfo(
        "Concluído",
        f"{total} arquivos gerados em:\n{os.path.abspath(outdir)}"
    )

# ===============================
# GUI mínima
# ===============================
def escolher_arquivo():
    entrada_arquivo.set(
        filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls *.xlsm")])
    )

def executar():
    if not entrada_arquivo.get():
        messagebox.showwarning("Aviso", "Selecione o arquivo de entrada.")
        return

    try:
        gerar_planilhas_por_cotacao(
            entrada_arquivo.get(),
            "../../outdir"
        )
    except Exception as e:
        messagebox.showerror("Erro", str(e))

# ===============================
# Interface
# ===============================
root = tk.Tk()
root.title("Gerador de Planilhas por Cotação")

entrada_arquivo = tk.StringVar()

tk.Label(root, text="Arquivo Excel de entrada:").grid(row=0, column=0, sticky="w")
tk.Entry(root, textvariable=entrada_arquivo, width=50).grid(row=0, column=1)
tk.Button(root, text="Procurar", command=escolher_arquivo).grid(row=0, column=2)

tk.Button(root, text="Gerar Arquivos", bg="green", fg="white", command=executar)\
    .grid(row=1, column=0, columnspan=3, pady=10)

root.mainloop()
