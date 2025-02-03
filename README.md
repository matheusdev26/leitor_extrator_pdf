# Leitor e Extrator de PDF 

## Descrição

Este projeto permite ao usuário escolher um arquivo PDF, extrair o texto dele e salvar o conteúdo extraído em um arquivo `.docx`. A interface gráfica foi criada utilizando a biblioteca `customtkinter`.

### Funcionalidades:
- **Escolher um arquivo PDF**: O usuário pode selecionar um arquivo PDF através de uma janela de diálogo.
- **Ler o conteúdo do PDF**: O programa lê o conteúdo do arquivo PDF e exibe uma prévia (primeiros 10000 caracteres) na interface gráfica.
- **Extrair e salvar em .docx**: O texto extraído é salvo em um arquivo `.docx` com o mesmo nome do PDF.
- **Abrir o arquivo .docx**: Após salvar, o programa pergunta se o usuário deseja abrir o arquivo gerado.

Este programa funciona em Windows, macOS e Linux.

## Requisitos

Este projeto requer as seguintes bibliotecas:

- **customtkinter**: Para a criação da interface gráfica.
- **PyPDF2**: Para ler e extrair o texto do PDF.
- **python-docx**: Para salvar o texto extraído em um arquivo `.docx`.

Você pode instalar as dependências com o seguinte comando:

```bash
pip install customtkinter PyPDF2 python-docx
