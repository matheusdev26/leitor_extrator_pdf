import customtkinter as ctk
from tkinter import filedialog
import PyPDF2
from docx import Document
from tkinter import messagebox
import os
import platform
import subprocess



def escolher_arquivo():
    global caminho_pdf
    caminho_pdf = filedialog.askopenfilename(
        title="Selecione um arquivo",
        filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")],
    )
    if caminho_pdf:
        label_arquivo.configure(text=f"Arquivo selecionado: {caminho_pdf}")

def ler_pdf():
    global texto_extraido
    if caminho_pdf:
        try:
            with open(caminho_pdf, "rb") as arquivo_pdf:
                leitor_pdf = PyPDF2.PdfReader(arquivo_pdf)
                texto = ""
                for pagina in leitor_pdf.pages:
                    texto += pagina.extract_text()
                texto_extraido = texto  # Armazena o texto extraído globalmente
                text_box.delete(1.0, "end")  # Limpa o conteúdo da caixa de texto
                text_box.insert("insert", texto[:10000])  # Exibe os primeiros 10000 caracteres
        except Exception as e:
            text_box.delete(1.0, "end")
            text_box.insert("insert", f"Erro ao ler o PDF: {e}")
    else:
        text_box.delete(1.0, "end")
        text_box.insert("insert", "Nenhum arquivo foi selecionado.")
        
def salvar_docx():
    global texto_extraido, caminho_pdf, caminho_doc
    
    if texto_extraido:
        # Cria um novo documento .docx
        doc = Document()

        # Adiciona o texto extraído ao doc como parágrafos
        doc.add_paragraph(texto_extraido)

        # Salva o doc com o mesmo nome do PDF, mas com extensão .docx
        caminho_doc = caminho_pdf.replace(".pdf", ".docx")
        doc.save(caminho_doc)

        label_arquivo.configure(text=f'O texto foi salvo em: {caminho_doc}')
        perguntar()
    else:
        label_arquivo.configure(text="Nenhum texto extraído para salvar.")

def perguntar():
    resposta = messagebox.askyesno("Confirmação", "Você deseja abrir o arquivo?")
    janela.withdraw()
    
    if resposta:
        abrir_arquivo(caminho_doc)


        
    else:
        janela.destroy()   
        


def abrir_arquivo(caminho_doc):
    try:
        if platform.system() == "Windows":  # Windows
            os.startfile(caminho_doc)
        elif platform.system() == "Darwin":  # macOS
            subprocess.run(["open", caminho_doc])
        else:  # Linux e outros Unix-like
            subprocess.run(["xdg-open", caminho_doc])
    except Exception as e:
        print(f"Erro ao abrir o arquivo: {e}")


# Configuração da janela principal
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

janela = ctk.CTk()
janela.title("Ler PDF")
janela.geometry("600x400")

# Botão para escolher o arquivo
botao_escolher = ctk.CTkButton(janela, text="Escolher PDF", command=escolher_arquivo)
botao_escolher.pack(pady=20)

# Botão para ler o conteúdo do PDF
botao_ler = ctk.CTkButton(janela, text="Ler PDF", command=ler_pdf)
botao_ler.pack(pady=20)

# Botão para extrair o texto e salvar em .docx
botao = ctk.CTkButton(janela, text="Extrair Texto", command=salvar_docx)
botao.pack(padx=20, pady=20)

# Label para exibir o caminho do arquivo selecionado
label_arquivo = ctk.CTkLabel(janela, text="Nenhum arquivo selecionado")
label_arquivo.pack(pady=10)

# Criação do TextBox com Scroll
text_box_frame = ctk.CTkFrame(janela)
text_box_frame.pack(pady=10, fill="both", expand=True)

# Caixa de texto com rolagem
text_box = ctk.CTkTextbox(text_box_frame, wrap="word", height=10)
text_box.pack(side="left", fill="both", expand=True)

# Barra de rolagem
scrollbar = ctk.CTkScrollbar(text_box_frame, command=text_box.yview)


# Inicia a aplicação
janela.mainloop()
