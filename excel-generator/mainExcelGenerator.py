import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
from datetime import datetime
import pandas as pd

# Função para ler o Excel de contatos
def carregar_contatos_excel(caminho_excel):
    contatos_dict = {}
    
    # Carrega o workbook
    wb = openpyxl.load_workbook(caminho_excel)
        
    # Seleciona a primeira planilha
    sheet = wb.active
    
    for row in sheet.iter_rows(min_row=2, values_only=True):  # Começa da segunda linha, assumindo que a primeira é cabeçalho
        if len(row) >= 4:  # Verifica se a linha tem pelo menos 4 colunas
            codigo, nomes, nome_contato, nome_grupo = row[:4]
            
            # Adiciona ao dicionário usando o código como chave
            contatos_dict[codigo] = {
                'contato': nome_contato,
                'grupo': nome_grupo
            }
    
    return contatos_dict

# Função para extrair informações do PDF
def extrair_informacoes_excel(caminho_excel_comparacao, contatos_dict):
    # Carregar dados do Excel para comparação
    df_comparacao = pd.read_excel(caminho_excel_comparacao)
    
    # Variaveis das colunas
    codigos = df_comparacao.iloc[:, 0]
    pessoas = df_comparacao.iloc[:, 1]

    # Criar o dicionário de dados onde cada chave (codigo) pode ter múltiplos registros
    dados = {}
    
    # Iterar sobre os códigos da coluna de codigos excel base
    for codigo_atual, pessoa in zip(codigos, pessoas):
        # Obter a data de vencimento como um objeto datetime
       
         
        # Verificar se a data de vencimento é maior ou igual à data atual
        
        # Verificar se o código está presente no dicionário de contatos
        if codigo_atual in contatos_dict:
            # Obter os contatos do dicionário
            contato_individual = contatos_dict[codigo_atual].get('contato', "")
            contato_grupo = contatos_dict[codigo_atual].get('grupo', "")
        else:
            contato_individual, contato_grupo = '', ''

        
        # Verificar se já existe uma lista associada ao código atual
        if codigo_atual not in dados:
            dados[codigo_atual] = []  # Se não, cria uma nova lista para armazenar dados do código
        
        # Adicionar ao dicionário: ao invés de sobrescrever, adicionamos à lista
        dados[codigo_atual].append({
            'Codigo': codigo_atual,
            'Empresa': pessoa,
            'Contato Onvio': contato_individual,
            'Grupo Onvio': contato_grupo
        })

    # Preparar a lista de linhas a serem retornadas
    linhas = []       
    for codigo, info_list in dados.items():
        for info in info_list:
            linhas.append(info)
    
    return linhas


# Função para gerar Excel a partir dos dados extraídos
def gerar_excel(dados, caminho_excel):
    df = pd.DataFrame(dados)
    df.to_excel(caminho_excel, index=False)
    print(f"Arquivo Excel criado: {caminho_excel}")

# Função para selecionar o arquivo PDF
def selecionar_excel_info():
    caminho_excel_info = filedialog.askopenfilename(
        title="Selecione o arquivo Informativo",
        filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
    )
    entrada_excel_base.delete(0, tk.END)
    entrada_excel_base.insert(0, caminho_excel_info)
    
# Função para selecionar o caminho para salvar o Excel
def selecionar_destino_excel():
    caminho_excel = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")),
        title="Salvar arquivo Excel"
    )
    entrada_excel.delete(0, tk.END)
    entrada_excel.insert(0, caminho_excel)

# Função para selecionar o caminho para salvar o Excel
def selecionar_lista_contatos():
    caminho_contatos = filedialog.askopenfilename(
        title="Selecione o arquivo Contatos",
        filetypes=(("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*"))
    )
    entrada_contatos.delete(0, tk.END)
    entrada_contatos.insert(0, caminho_contatos)

# Função para processar o PDF e gerar o Excel
def processar():

    caminho_excel_base = entrada_excel_base.get()
    caminho_excel = entrada_excel.get()
    caminho_contatos = entrada_contatos.get() #Excel Lista de contatos Onvio

    if not caminho_excel_base or not caminho_excel or not caminho_contatos:
        messagebox.showwarning("Erro", "Por favor, selecione o arquivo PDF e o local para salvar o Excel.")
        return
    
    try:
        # Carrega o Excel com a lista de contatos
        contatos_dict = carregar_contatos_excel(caminho_contatos)
        
        linhas_extraidas = extrair_informacoes_excel(caminho_excel_base, contatos_dict)
        if linhas_extraidas:
            gerar_excel(linhas_extraidas, caminho_excel)
            messagebox.showinfo("Sucesso", "O arquivo Excel foi gerado com sucesso!")
        else:
            messagebox.showwarning("Erro", "Nenhum dado foi extraído do PDF.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar: {e}")    


def main():
    global entrada_excel_base, entrada_excel, entrada_contatos 
     
    # Interface gráfica com Tkinter
    janela = tk.Tk()
    janela.title("Gerador do Excel Renovação de Contratos")
    janela.geometry("500x400")

    # Campo para o caminho do PDF
    lbl_pdf = tk.Label(janela, text="Selecione o Excel Base:")
    lbl_pdf.pack(pady=5)
    entrada_excel_base = tk.Entry(janela, width=50)
    entrada_excel_base.pack(pady=5)
    btn_pdf = tk.Button(janela, text="Selecionar Excel Base", command=selecionar_excel_info)
    btn_pdf.pack(pady=5)

    # Campo para o caminho do Excel com a lista de contatos
    lbl_contatos = tk.Label(janela, text="Selecione o Excel de Contatos:")
    lbl_contatos.pack(pady=5)
    entrada_contatos = tk.Entry(janela, width=50)
    entrada_contatos.pack(pady=5)
    btn_contatos = tk.Button(janela, text="Selecionar Excel de Contatos", command=selecionar_lista_contatos)
    btn_contatos.pack(pady=5)

    # Campo para o caminho do Excel
    lbl_excel = tk.Label(janela, text="Selecione o destino do arquivo Excel:")
    lbl_excel.pack(pady=5)
    entrada_excel = tk.Entry(janela, width=50)
    entrada_excel.pack(pady=5)
    btn_excel = tk.Button(janela, text="Salvar Excel como", command=selecionar_destino_excel)
    btn_excel.pack(pady=5)

    # Botão para processar o PDF e gerar o Excel
    btn_processar = tk.Button(janela, text="Gerar Excel", command=processar)
    btn_processar.pack(pady=10)

    # Iniciar a janela
    janela.mainloop()

if __name__ == '__main__':
    main()