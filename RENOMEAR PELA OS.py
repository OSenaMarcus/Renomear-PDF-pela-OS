import os
import openpyxl
import tkinter as tk
from tkinter import filedialog, ttk
from tkinter import messagebox  # Importar o módulo messagebox
from pdf2image import convert_from_path
import pytesseract

# Função para selecionar a pasta com os PDFs
def selecionar_pasta_pdfs():
    pasta_pdfs = filedialog.askdirectory()
    return pasta_pdfs

# Função para selecionar a planilha do Excel
def selecionar_planilha_excel():
    planilha_excel = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    return planilha_excel

# Função para extrair o texto do PDF usando OCR
def extrair_texto_pdf(pdf_path):
    try:
        # Converter o PDF em imagens
        images = convert_from_path(pdf_path)
        
        text = ""
        for image in images:
            # Extrair o texto da imagem usando OCR
            text += pytesseract.image_to_string(image, lang='por')
        
        return text
    except Exception as e:
        print(f"Erro ao extrair texto do PDF: {str(e)}")
        return None

# Função para processar os PDFs
def processar_pdfs():
    pasta_pdfs = selecionar_pasta_pdfs()
    planilha_excel = selecionar_planilha_excel()
    
    if pasta_pdfs and planilha_excel:
        print(f"Pasta PDFs selecionada: {pasta_pdfs}")
        print(f"Planilha Excel selecionada: {planilha_excel}")
        
        # Carregar a planilha do Excel
        workbook = openpyxl.load_workbook(planilha_excel)
        sheet = workbook.active
        
        # Obter os valores das colunas OS, placa, ait, lote lummon e cliente
        os_values = [cell.value for cell in sheet['C'][1:]]
        placa_values = [cell.value for cell in sheet['F'][1:]]
        ait_values = [cell.value for cell in sheet['E'][1:]]
        lote_lummon_values = [cell.value for cell in sheet['D'][1:]]
        cliente_values = [cell.value for cell in sheet['B'][1:]]
        
        # Dicionário para armazenar os lotes e seus respectivos PDFs
        lotes = {}
        
        # Dicionário para armazenar os clientes e seus respectivos lotes
        clientes = {}
        
        # Iterar sobre os PDFs na pasta selecionada
        for filename in os.listdir(pasta_pdfs):
            if filename.endswith(".pdf"):
                pdf_path = os.path.join(pasta_pdfs, filename)
                
                # Extrair o texto do PDF usando OCR
                pdf_text = extrair_texto_pdf(pdf_path)
                
                if pdf_text:
                    # Verificar se o texto do PDF contém algum valor da coluna OS
                    for i, os_value in enumerate(os_values):
                        if str(os_value) in pdf_text:
                            # Verificar se o valor da coluna CLIENTE é "Kintomobility"
                            if str(cliente_values[i]) == "Kintomobility":
                                # Renomear o PDF com o valor da coluna PLACA + AIT + "_C" + "PG"
                                new_filename = f"{placa_values[i]} {ait_values[i]}_C PG.pdf"
                            else:
                                # Verificar se os valores das colunas PLACA e AIT são "-"
                                if str(placa_values[i]) == "-" and str(ait_values[i]) == "-":
                                    # Renomear o PDF somente com o valor da coluna OS + ' PG'
                                    new_filename = f"{os_value} PG.pdf"
                                else:
                                    # Verificar se o valor da coluna AIT é "-"
                                    if str(ait_values[i]) == "-":
                                        # Renomear o PDF com o valor da coluna OS + coluna placa + 'PG'
                                        new_filename = f"{os_value} {placa_values[i]} PG.pdf"
                                    else:
                                        # Renomear o PDF com o valor da coluna placa + coluna ait + 'PG'
                                        new_filename = f"{placa_values[i]} {ait_values[i]} PG.pdf"
                            
                            new_path = os.path.join(pasta_pdfs, new_filename)
                            
                            # Verificar se o arquivo de destino já existe
                            if os.path.exists(new_path):
                                # Adicionar um sufixo numérico ao nome do arquivo
                                counter = 1
                                while True:
                                    new_filename_with_suffix = f"{os.path.splitext(new_filename)[0]}_{counter}.pdf"
                                    new_path_with_suffix = os.path.join(pasta_pdfs, new_filename_with_suffix)
                                    if not os.path.exists(new_path_with_suffix):
                                        new_filename = new_filename_with_suffix
                                        new_path = new_path_with_suffix
                                        break
                                    counter += 1
                            
                            os.rename(pdf_path, new_path)
                            print(f"PDF renomeado: {new_filename}")
                            
                            # Obter o valor do lote lummon correspondente
                            lote_lummon = lote_lummon_values[i]
                            
                            # Adicionar o PDF ao dicionário de lotes
                            if lote_lummon not in lotes:
                                lotes[lote_lummon] = []
                            lotes[lote_lummon].append(new_filename)
                            
                            # Obter o valor do cliente correspondente
                            cliente = cliente_values[i]
                            
                            # Adicionar o lote ao dicionário de clientes
                            if cliente not in clientes:
                                clientes[cliente] = []
                            if lote_lummon not in clientes[cliente]:
                                clientes[cliente].append(lote_lummon)
                            
                            break
        
        # Criar pastas para cada lote e mover os PDFs correspondentes
        for lote, pdfs in lotes.items():
            lote_folder = os.path.join(pasta_pdfs, str(lote))
            os.makedirs(lote_folder, exist_ok=True)
            
            for pdf in pdfs:
                pdf_path = os.path.join(pasta_pdfs, pdf)
                new_path = os.path.join(lote_folder, pdf)
                os.rename(pdf_path, new_path)
                print(f"PDF movido para o lote {lote}: {pdf}")
        
        # Criar pastas para cada cliente e mover os lotes correspondentes
        for cliente, lotes_cliente in clientes.items():
            cliente_folder = os.path.join(pasta_pdfs, str(cliente))
            os.makedirs(cliente_folder, exist_ok=True)
            
            for lote in lotes_cliente:
                lote_folder = os.path.join(pasta_pdfs, str(lote))
                new_lote_folder = os.path.join(cliente_folder, str(lote))
                
                # Verificar se o lote existe antes de mover
                if os.path.exists(lote_folder):
                    os.rename(lote_folder, new_lote_folder)
                    print(f"Lote {lote} movido para a pasta do cliente {cliente}")
                else:
                    print(f"O lote {lote} não foi encontrado para o cliente {cliente}")
        
        # Exibir mensagem de conclusão
        messagebox.showinfo("Processamento Concluído", "O processamento dos PDFs foi concluído com sucesso!")
        
        print("Processamento concluído.")
    else:
        print("Por favor, selecione a pasta com os PDFs e a planilha do Excel.")

def exibir_ajuda():
    mensagem_ajuda = """
    Como usar o programa "Renomear PDFs":

    1. Clique no botão "Processar PDFs" para iniciar o processo.
    2. Selecione a pasta onde estão localizados os arquivos PDF que deseja renomear.
    3. Selecione a planilha do Excel que contém as informações para renomear os PDFs.
    4. O programa irá processar os PDFs na pasta selecionada.
    5. Os PDFs serão renomeados de acordo com as seguintes regras:
       - Se os valores das colunas "PLACA" e "AIT" forem "-", o PDF será renomeado somente com o valor da coluna "OS" + "PG".
       - Se apenas o valor da coluna "AIT" for "-", o PDF será renomeado com o valor da coluna "OS" + coluna "PLACA" + "PG".
       - Caso contrário, o PDF será renomeado com o valor da coluna "PLACA" + coluna "AIT" + "PG".
    6. Após a renomeação, os PDFs serão movidos para pastas correspondentes aos valores da coluna "LOTE LUMMON".
    7. O processamento será concluído e uma mensagem de conclusão será exibida.

    Observações:
    - Certifique-se de que a planilha do Excel esteja no formato correto, com as colunas "OS", "PLACA", "AIT" e "LOTE LUMMON".
    - Os arquivos PDF devem estar na pasta selecionada para serem processados corretamente.
    - O programa utiliza o OCR para extrair o texto dos PDFs, portanto, a qualidade da renomeação pode variar dependendo da qualidade dos PDFs.
    """
    messagebox.showinfo("Ajuda", mensagem_ajuda)

# Criar a janela principal
window = tk.Tk()
window.title("Renomear PDFs")

# Definir o tamanho da janela e desabilitar a opção de maximizar
window.geometry("225x200")  # Largura x Altura
window.resizable(False, False)  # Desabilitar redimensionamento horizontal e vertical

# Criar um estilo personalizado para os botões
style = ttk.Style()
style.configure("Red.TButton",
                padding=5,  # Diminuir o valor do padding para reduzir o tamanho dos botões
                relief="flat",
                background="#ff0000",
                foreground="black",
                font=("Arial", 10, "bold"))  # Reduzir o tamanho da fonte para 10
style.map("Red.TButton",
          background=[("active", "#cc0000")],
          foreground=[("active", "black")])

# Criar o botão "Processar PDFs" usando o estilo personalizado
button_processar = ttk.Button(window, text="Processar PDFs", style="Red.TButton", command=processar_pdfs)
button_processar.pack(pady=20)  # Adicionar espaçamento vertical

# Criar o botão "Ajuda" usando o mesmo estilo personalizado
button_ajuda = ttk.Button(window, text="Ajuda", style="Red.TButton", command=exibir_ajuda)
button_ajuda.pack(pady=10)  # Adicionar espaçamento vertical

# Criar o rodapé com o texto "Desenvolvido por: Marcus Sena"
footer_frame = ttk.Frame(window)
footer_frame.pack(side="bottom", fill="x")

footer_label = ttk.Label(footer_frame, text="Desenvolvido por: Marcus Sena", font=("Arial", 10))
footer_label.pack(side="right", padx=10, pady=5)

# Iniciar o loop principal da janela
window.mainloop()
