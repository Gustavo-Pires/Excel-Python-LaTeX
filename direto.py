import pandas as pd
from fpdf import FPDF

# --- CONFIGURAÇÕES ---
excel_file_path = "dados.xlsx"
timbrado_path = "timbrado.jpg"
imagem_path = "imagem.png"
fonte_sofia_path = "sofiapro-regular.ttf"  # fonte baixada




# --- LENDO OS DADOS DO EXCEL ---
df = pd.read_excel(excel_file_path, usecols="B", nrows=21, header=None)
dados = [str(v) if pd.notna(v) else "" for v in df.iloc[:, 0].tolist()]

# Gera o nome do PDF a partir de dados[0]
pdf_output_path = f"{dados[0]}.pdf"

# --- CLASSE PDF ---
class PDF(FPDF):
    def header(self):
        self.image(timbrado_path, x=0, y=0, w=self.w, h=self.h)

    def chapter_body(self, body):
        self.set_font('SofiaPro', '', 12)
        self.multi_cell(0, 8, body)
        self.ln()

    def colored_table(self, title, rows):
        self.set_fill_color(47, 63, 79)
        self.set_text_color(255, 255, 255)
        self.set_font('SofiaPro', '', 12)
        self.cell(0, 8, title, border=1, ln=True, align='C', fill=True)

        self.set_font('SofiaPro', '', 11)
        self.set_text_color(0, 0, 0)
        col1_width = 70
        col2_width = self.w - col1_width - 20

        fill = False
        for key, value in rows:
            self.set_fill_color(230, 230, 230) if fill else self.set_fill_color(255, 255, 255)
            self.cell(col1_width, 8, key, border=1, fill=fill)
            self.cell(col2_width, 8, value, border=1, ln=True, fill=fill)
            fill = not fill

# --- CRIANDO PDF ---
pdf = PDF()
pdf.add_font('SofiaPro', '', fonte_sofia_path, uni=True)  # registra fonte

pdf.add_page()

# Título
pdf.set_font('SofiaPro', '', 18)
pdf.cell(0, 10, f"{dados[0]} -- Descrição do Item", ln=True)

# Descrição
pdf.set_font('SofiaPro', '', 14)
pdf.cell(0, 10, "Descrição e Funcionalidade", ln=True)
pdf.set_font('SofiaPro', '', 12)
pdf.multi_cell(0, 8, dados[1])
pdf.ln()

# Imagem centralizada
pdf.image(imagem_path, x=(pdf.w - 50) / 2, w=50)
pdf.ln(5)

# Tabela 1
pdf.colored_table("Informações Gerais", [
    ("Código", dados[2]),
    ("Nome", dados[3]),
    ("Categoria", dados[4]),
    ("Fornecedor", dados[5]),
    ("NCM", dados[6]),
    ("Equip. Relacionados", dados[7])
])

pdf.ln(5)

# Tabela 2
pdf.colored_table("Características Físicas", [
    ("Altura", dados[8]),
    ("Largura", dados[9]),
    ("Profundidade", dados[10]),
    ("Diâmetro", dados[11]),
    ("Peso aproximado", dados[12]),
])

# Tabela 3
pdf.colored_table("Características Técnicas", [
    ("Material principal", dados[13]),
    ("Cor predominante", dados[14]),
    ("Tensão de operação", dados[15]),
    ("Corrente nominal", dados[16]),
    ("Pressão de trabalho", dados[17]),
    ("Conexões", dados[18]),
    ("Grau de proteção IP", dados[19]),
    ("Info. adicionais", dados[20]),
])

# Salva PDF
pdf.output(pdf_output_path)
print(f"PDF gerado em: {pdf_output_path}")
