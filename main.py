"""
Projeto: Leitor de PDF com Interface Gráfica

Descrição:
Este programa permite ao usuário escolher um arquivo PDF, extrair o texto do PDF e salvar o conteúdo extraído em um arquivo .docx. 
A interface gráfica é construída utilizando a biblioteca `customtkinter` para criar uma experiência mais moderna em comparação com o tkinter padrão.

Requisitos:
- customtkinter
- PyPDF2
- python-docx

Funcionalidades:
1. Escolher um arquivo PDF através de uma janela de diálogo.
2. Ler o conteúdo do arquivo PDF e exibir os primeiros 1000 caracteres na interface.
3. Salvar o texto extraído em um novo arquivo .docx.
4. Perguntar se o usuário deseja abrir o arquivo .docx depois de salvo.
5. Compatível com Windows, macOS e Linux para abrir o arquivo gerado.

Autor: Matheus Souza
Data: Fevereiro de 2025
"""

import customtkinter as ctk
from tkinter import filedialog
import PyPDF2
from docx import Document
from tkinter import messagebox
import os
import platform
import subprocess

# Função para escolher um arquivo PDF
def escolher_arquivo():
    """
    Abre uma janela de diálogo para o usuário escolher um arquivo PDF.
    O caminho do arquivo selecionado é armazenado na variável global `caminho_pdf`.
    """
    global caminho_pdf
    caminho_pdf = filedialog.askopenfilename(
        title="Selecione um arquivo",
        filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")],
    )
    if caminho_pdf:
        # Atualiza o label para mostrar o caminho do arquivo selecionado
        label_arquivo.configure(text=f"Arquivo selecionado: {caminho_pdf}")

# Função para ler o conteúdo do PDF
def ler_pdf():
    """
    Lê o conteúdo do arquivo PDF selecionado e extrai o texto.
    O texto extraído é exibido parcialmente (primeiros 1000 caracteres) em uma caixa de texto.
    Se ocorrer um erro durante a leitura, a mensagem de erro é exibida na caixa de texto.
    """
    global texto_extraido
    if caminho_pdf:
        try:
            # Abre o PDF e lê seu conteúdo
            with open(caminho_pdf, "rb") as arquivo_pdf:
                leitor_pdf = PyPDF2.PdfReader(arquivo_pdf)
                texto = ""
                for pagina in leitor_pdf.pages:
                    texto += pagina.extract_text()
                texto_extraido = texto  # Armazena o texto extraído globalmente
                text_box.delete(1.0, "end")  # Limpa o conteúdo da caixa de texto
                text_box.insert("insert", texto[:1000])  # Exibe os primeiros 1000 caracteres
        except Exception as e:
            # Se houver um erro ao ler o PDF, exibe a mensagem de erro
            text_box.delete(1.0, "end")
            text_box.insert("insert", f"Erro ao ler o PDF: {e}")
    else:
        # Caso nenhum arquivo tenha sido selecionado
        text_box.delete(1.0, "end")
        text_box.insert("insert", "Nenhum arquivo foi selecionado.")

# Função para salvar o texto extraído em um arquivo .docx
def salvar_docx():
    """
    Salva o texto extraído do PDF em um arquivo .docx com o mesmo nome do arquivo PDF.
    Se não houver texto extraído, exibe uma mensagem de erro.
    """
    global texto_extraido, caminho_pdf, caminho_doc
    
    if texto_extraido:
        # Cria um novo documento .docx
        doc = Document()

        # Adiciona o texto extraído ao documento como parágrafos
        doc.add_paragraph(texto_extraido)

        # Salva o documento .docx com o mesmo nome do arquivo PDF
        caminho_doc = caminho_pdf.replace(".pdf", ".docx")
        doc.save(caminho_doc)

        label_arquivo.configure(text=f'O texto foi salvo em: {caminho_doc}')
        perguntar()  # Pergunta se o usuário deseja abrir o arquivo após o salvamento
    else:
        label_arquivo.configure(text="Nenhum texto extraído para salvar.")

# Função para perguntar ao usuário se deseja abrir o arquivo .docx
def perguntar():
    """
    Exibe uma caixa de mensagem perguntando se o usuário deseja abrir o arquivo .docx após o salvamento.
    Se o usuário escolher 'Sim', o arquivo será aberto. Caso contrário, o programa será fechado.
    """
    resposta = messagebox.askyesno("Confirmação", "Você deseja abrir o arquivo?")
    janela.withdraw()
    
    if resposta:
        # Chama a função para abrir o arquivo
        abrir_arquivo(caminho_doc)
    else:
        janela.destroy()   # Fecha a janela principal se o usuário não quiser abrir o arquivo

# Função para abrir o arquivo .docx
def abrir_arquivo(caminho_doc):
    """
    Abre o arquivo .docx salvo, dependendo do sistema operacional.
    - No Windows, usa `os.startfile()`.
    - No macOS, usa o comando `open`.
    - No Linux, usa o comando `xdg-open`.
    """
    try:
        if platform.system() == "Windows":  # Windows
            os.startfile(caminho_doc)
        elif platform.system() == "Darwin":  # macOS
            subprocess.run(["open", caminho_doc])
        else:  # Linux e outros Unix-like
            subprocess.run(["xdg-open", caminho_doc])
    except Exception as e:
        print(f"Erro ao abrir o arquivo: {e}")

# Configuração da janela principal usando customtkinter
ctk.set_appearance_mode("System")  # Configura o modo de aparência do sistema
ctk.set_default_color_theme("blue")  # Define o tema de cores como "blue"

# Criação da janela principal
janela = ctk.CTk()
janela.title("Ler PDF")
janela.geometry("600x400")

# Botões de interação com o usuário
botao_escolher = ctk.CTkButton(janela, text="Escolher PDF", command=escolher_arquivo)
botao_escolher.pack(pady=20)

botao_ler = ctk.CTkButton(janela, text="Ler PDF", command=ler_pdf)
botao_ler.pack(pady=20)

botao = ctk.CTkButton(janela, text="Extrair Texto", command=salvar_docx)
botao.pack(padx=20, pady=20)

# Label para exibir o caminho do arquivo selecionado
label_arquivo = ctk.CTkLabel(janela, text="Nenhum arquivo selecionado")
label_arquivo.pack(pady=10)

# Caixa de texto para exibir parte do conteúdo extraído
text_box_frame = ctk.CTkFrame(janela)
text_box_frame.pack(pady=10, fill="both", expand=True)

# Caixa de texto com rolagem
text_box = ctk.CTkTextbox(text_box_frame, wrap="word", height=10)
text_box.pack(side="left", fill="both", expand=True)

# Barra de rolagem para a caixa de texto
scrollbar = ctk.CTkScrollbar(text_box_frame, command=text_box.yview)
scrollbar.pack(side="right", fill="y")

# Inicia a aplicação
janela.mainloop()
