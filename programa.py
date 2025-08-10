import pandas as pd
import re
import subprocess
import os

# Caminhos dos arquivos
excel_path = "dados.xlsx"
latex_dir = "LaTeX"
latex_file = "main.tex"
latex_path = os.path.join(latex_dir, latex_file)

# 1. Lê os dados do Excel
df = pd.read_excel(excel_path, sheet_name=0, header=None)
df = df.dropna().reset_index(drop=True)

# 2. Cria nova tabela LaTeX
nova_tabela = []
for _, row in df.iterrows():
    campo, valor = row[0], row[1]
    linha = f"\\multicolumn{{1}}{{|l|}}{{{campo}}} & {valor} \\\\ \\hline"
    nova_tabela.append(linha)
nova_tabela_tex = "\n    ".join(nova_tabela)

# 3. Lê o LaTeX original
with open(latex_path, "r", encoding="utf-8") as f:
    conteudo = f.read()

# 4. Regex para encontrar o bloco da tabela
padrao = re.compile(
    r"(\\multicolumn\{2\}\{\|c\|\}\{\\cellcolor\[HTML\]\{002060\}\{\\color\[HTML\]\{FFFFFF\} Informacoes\s+Gerais\}\} \\\\ \\hline\n)(.*?)(\\end\{tabular\})",
    flags=re.DOTALL
)

# 5. Substitui a parte da tabela
def substitui_tabela(match):
    inicio = match.group(1)
    fim = match.group(3)
    return inicio + nova_tabela_tex + "\n    " + fim

novo_conteudo = padrao.sub(substitui_tabela, conteudo)

# 6. Salva o arquivo modificado
with open(latex_path, "w", encoding="utf-8") as f:
    f.write(novo_conteudo)

print("Tabela LaTeX atualizada com sucesso.")

# 7. Compila o arquivo LaTeX para PDF com pdflatex
try:
    resultado = subprocess.run(
        ["pdflatex", "-interaction=nonstopmode", latex_file],
        cwd=latex_dir,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        check=True
    )
    print("✅ PDF compilado com sucesso.")
except subprocess.CalledProcessError as e:
    print("❌ Erro ao compilar o LaTeX:")
    print(e.stdout)
    print(e.stderr)