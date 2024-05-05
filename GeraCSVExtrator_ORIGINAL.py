import tkinter as tk
from tkinter import messagebox, filedialog
import pyodbc
import csv

#chamando a funções 
def generate_csv():
    server = server_entry.get()
    database = database_entry.get()
    procedure = procedure_entry.get()
    period_start = period_start_entry.get() 
    period_end = period_end_entry.get()
    parameters = parameters_entry.get()
    save_path = save_path_entry.get()
    include_cnpj = include_cnpj_var.get() 

    try:
        # Estabelecer conexão com o banco de dados
        conn = pyodbc.connect(f"Driver={{SQL Server}};Server={server};Database={database};Trusted_Connection=yes;")
        cursor = conn.cursor()

        try:
            # Encontrando a tabela "tb_entradas" ou qualquer tabela que contenha a palavra "entradas" (depois ver se tem mais tb com outro nomes nas filiais)
            cursor.execute("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME LIKE '%entradas%'")
            tables = [row.TABLE_NAME for row in cursor.fetchall()]

            if not tables:
                messagebox.showerror("Erro", "Nenhuma tabela contendo 'entradas' foi encontrada.")
                return

            try:
                # Tentar executar a consulta usando a tabela "tb_entradas"
                cursor.execute(f"SELECT DISTINCT CGC FROM {tables[0]}")
            except pyodbc.ProgrammingError:
                # Se a consulta com a tabela "tb_entradas" falhar, tentar com a primeira tabela que contenha "entradas"
                cursor.execute(f"SELECT DISTINCT CNPJ_FILIAL AS CGC FROM {tables[0]}")

            cnpjs_filiais = [row.CGC for row in cursor.fetchall()]

            with open(save_path, 'w', newline='') as csvfile:
                csv_writer = csv.writer(csvfile, delimiter=';')

                # Extrair anos do periodo inicial e 0 periodo final 
                start_year = int(period_start.split("_")[-1])
                end_year = int(period_end.split("_")[-1])

                # Gerar os arquivos em CSV 
                for year in range(start_year, end_year + 1):
                    for CGC in cnpjs_filiais:
                        for i in range(1, 13):
                            filename = f"TXT_{CGC}_{i:02d}_{year}.txt"
                            # parte do CNPJ selecionavel
                            if include_cnpj:
                                row_data = [filename, excluir_var.get(), incluir_var.get(), server, database, procedure, CGC, i, year]
                            else:
                                row_data = [filename, excluir_var.get(), incluir_var.get(), server, database, procedure, i, year]
                            if parameters:
                                # Verificar se os parâmetros são numéricos mesmoo
                                if parameters.isdigit():
                                    row_data.append(int(parameters))
                                else:
                                    row_data.append(parameters)  #Aqui não tem adição de aspas
                            csv_writer.writerow(row_data)
            messagebox.showinfo("Sucesso", "Arquivo CSV gerado com sucesso!")

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

    finally:
        cursor.close()
        conn.close()


def browse_path():
    path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("Arquivo CSV", "*.csv")])
    if path:
        save_path_entry.delete(0, tk.END)
        save_path_entry.insert(0, path)

# Criar a janela principal (depois mudar a cor)
root = tk.Tk()
root.title("Gerador de Arquivo CSV para a ferramenta Extrator")
root.geometry("700x300")


# Servidor e a base
server_label = tk.Label(root, text="Servidor:")
server_label.grid(row=0, column=0, sticky="w")
server_entry = tk.Entry(root)
server_entry.grid(row=0, column=1)

database_label = tk.Label(root, text="Banco:")
database_label.grid(row=1, column=0, sticky="w")
database_entry = tk.Entry(root)
database_entry.grid(row=1, column=1)

# A Procedure
procedure_label = tk.Label(root, text="Procedure:")
procedure_label.grid(row=2, column=0, sticky="w")
procedure_entry = tk.Entry(root)
procedure_entry.grid(row=2, column=1, columnspan=3, sticky="we")
procedure_entry.insert(0, "")

# Colocando periodo
year_period_label = tk.Label(root, text="Período de (Somente Ano):")
year_period_label.grid(row=3, column=0, sticky="w")
period_start_entry = tk.Entry(root)
period_start_entry.grid(row=3, column=1)
year_period_label = tk.Label(root, text="Até (Somente Ano):")
year_period_label.grid(row=3, column=2, sticky="w")
period_end_entry = tk.Entry(root)
period_end_entry.grid(row=3, column=3)

# os parametros caso existaa
parameters_label = tk.Label(root, text="Parâmetros (se existir):")
parameters_label.grid(row=4, column=0, sticky="w")
parameters_entry = tk.Entry(root)
parameters_entry.grid(row=4, column=1)

# salvando o caminho
save_path_label = tk.Label(root, text="Caminho para salvar o arquivo .CSV:")
save_path_label.grid(row=5, column=0, sticky="w")
save_path_entry = tk.Entry(root)
save_path_entry.grid(row=5, column=1)
browse_button = tk.Button(root, text="Pasta de destino", command=browse_path)
browse_button.grid(row=5, column=2)

# Inclusão do CNPJ/filial se quiser, ele tika
include_cnpj_var = tk.IntVar(value=1)
include_cnpj_checkbutton = tk.Checkbutton(root, text="Incluir a filial como parametro?", variable=include_cnpj_var)
include_cnpj_checkbutton.grid(row=6, column=0, columnspan=4, sticky="w")

# Excluir se existir parametro da ferramenta extrator (opção qda ferramenta)
excluir_var = tk.IntVar()
excluir_checkbutton = tk.Checkbutton(root, text="Excluir se existir?", variable=excluir_var)
excluir_checkbutton.grid(row=7, column=0, columnspan=2, sticky="w")

incluir_var = tk.IntVar()
incluir_checkbutton = tk.Checkbutton(root, text="Incluir cabeçalho?", variable=incluir_var)
incluir_checkbutton.grid(row=7, column=2, columnspan=2, sticky="e")

# Generação do botão (depois deixar mais apresentavel)
generate_button = tk.Button(root, text="Gerar o .CSV", font="Arial 12 bold", fg="blue",  command=generate_csv)
generate_button.grid(row=8, column=0, columnspan=4)

# Aviso que vai no final da ordem
note_label = tk.Label(root, text="OBS: Na procedure do banco de dados, se possível, coloque como parâmetro o CNPJ, MES e ANO.", font=("Helvetica", 10, "bold"))
note_label.grid(row=9, column=0, columnspan=4, pady=(10, 0))

root.mainloop()
