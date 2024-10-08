import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

# Função para somar os valores da base de ICMS e ICMS destacado por alíquota
def calcular_somas_por_aliquota(df):
    return df.groupby('Alíquota de ICMS').agg({
        'Base de ICMS': 'sum',
        'Valor do ICMS': 'sum'
    }).reset_index()

def carregar_arquivo():
    file_path = filedialog.askopenfilename(title="Selecione o arquivo CSV", filetypes=[("CSV Files", "*.csv")])
    if file_path:
        try:
            # Atualização do parâmetro de leitura do CSV
            df = pd.read_csv(file_path, sep=';', encoding='latin1', on_bad_lines='skip')

            # Verificar se as colunas necessárias existem
            colunas_necessarias = ['Alíquota de ICMS', 'CST ICMS', 'CFOP', 'Base de ICMS', 'Valor do ICMS']
            colunas_faltantes = [col for col in colunas_necessarias if col not in df.columns]
            if colunas_faltantes:
                messagebox.showerror("Erro", f"As seguintes colunas estão faltando no arquivo CSV: {', '.join(colunas_faltantes)}")
                return

            # Corrigir a formatação da coluna CST ICMS
            df['CST ICMS'] = df['CST ICMS'].str.replace('="', '').str.replace('"', '').str.strip()

            # Verifique os valores únicos da coluna CST ICMS após a correção
            print("Valores corrigidos em 'CST ICMS':", df['CST ICMS'].unique())

            # Converter 'Alíquota de ICMS' para float
            df['Alíquota de ICMS'] = df['Alíquota de ICMS'].replace(',', '.', regex=True).astype(float)

            # Filtrar as alíquotas desejadas
            aliquotas_desejadas = [4, 10, 12, 17, 25]
            df_filtrado = df[df['Alíquota de ICMS'].isin(aliquotas_desejadas)]

            # Filtrar as CST desejadas
            cst_desejadas = ['1/00', '1/10', '1/20', '1/51', '1/90']
            df_filtrado = df_filtrado.copy()
            df_filtrado['CST ICMS'] = df_filtrado['CST ICMS'].astype(str).str.strip()
            df_filtrado = df_filtrado[df_filtrado['CST ICMS'].isin(cst_desejadas)]

            # Verificar se há dados após a filtragem
            if df_filtrado.empty:
                messagebox.showinfo("Informação", "Nenhum dado corresponde aos filtros aplicados.")
                return

            # Separar operações interestaduais e estaduais
            df_filtrado['CFOP'] = df_filtrado['CFOP'].astype(str)
            df_interestaduais = df_filtrado[df_filtrado['CFOP'].str.startswith('6')]
            df_estaduais = df_filtrado[df_filtrado['CFOP'].str.startswith('5')]

            # Corrigir valores para operações interestaduais
            df_interestaduais['Base de ICMS'] = df_interestaduais['Base de ICMS'].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
            df_interestaduais['Valor do ICMS'] = df_interestaduais['Valor do ICMS'].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
            df_interestaduais['Base de ICMS'] = pd.to_numeric(df_interestaduais['Base de ICMS'], errors='coerce')
            df_interestaduais['Valor do ICMS'] = pd.to_numeric(df_interestaduais['Valor do ICMS'], errors='coerce')
            df_interestaduais = df_interestaduais.dropna(subset=['Base de ICMS', 'Valor do ICMS'])

            # Corrigir valores para operações estaduais
            df_estaduais['Base de ICMS'] = df_estaduais['Base de ICMS'].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
            df_estaduais['Valor do ICMS'] = df_estaduais['Valor do ICMS'].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
            df_estaduais['Base de ICMS'] = pd.to_numeric(df_estaduais['Base de ICMS'], errors='coerce')
            df_estaduais['Valor do ICMS'] = pd.to_numeric(df_estaduais['Valor do ICMS'], errors='coerce')
            df_estaduais = df_estaduais.dropna(subset=['Base de ICMS', 'Valor do ICMS'])

            # Calcular as somas por alíquota
            somas_interestaduais = calcular_somas_por_aliquota(df_interestaduais)
            somas_estaduais = calcular_somas_por_aliquota(df_estaduais)

             # Verificar se há resultados a serem exibidos
            if somas_interestaduais.empty and somas_estaduais.empty:
                messagebox.showinfo("Informação", "Nenhum resultado para exibir após o processamento.")
                return
            
            # Formatar valores no padrão brasileiro
            somas_interestaduais['Base de ICMS'] = somas_interestaduais['Base de ICMS'].apply(formatar_valores_brasileiros)
            somas_interestaduais['Valor do ICMS'] = somas_interestaduais['Valor do ICMS'].apply(formatar_valores_brasileiros)
            somas_estaduais['Base de ICMS'] = somas_estaduais['Base de ICMS'].apply(formatar_valores_brasileiros)
            somas_estaduais['Valor do ICMS'] = somas_estaduais['Valor do ICMS'].apply(formatar_valores_brasileiros)

             # Exibir resultados
            exibir_resultados(somas_interestaduais, somas_estaduais)
        except Exception as e:
            import traceback
            traceback_str = traceback.format_exc()
            messagebox.showerror("Erro", f"Ocorreu um erro ao processar o arquivo: {e}\n\n{traceback_str}")
            
            
# Função para formatar valores no padrão brasileiro
def formatar_valores_brasileiros(valor):
    return f"{valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")           

# Função para exibir os resultados em uma nova janela
def exibir_resultados(somas_interestaduais, somas_estaduais):
    resultado_window = tk.Toplevel()
    resultado_window.title("Resultados - Somas por Alíquota")

    # Configuração do notebook para separar as abas
    notebook = ttk.Notebook(resultado_window)
    notebook.pack(fill='both', expand=True)

    # Verificar e adicionar abas apenas se houver dados
    if not somas_interestaduais.empty:
        # Frame para operações interestaduais
        frame_interestaduais = ttk.Frame(notebook)
        notebook.add(frame_interestaduais, text="Operações Interestaduais")

        tree_interestaduais = ttk.Treeview(frame_interestaduais, columns=("Alíquota", "Base de ICMS", "Valor de ICMS"), show="headings")
        tree_interestaduais.heading("Alíquota", text="Alíquota de ICMS")
        tree_interestaduais.heading("Base de ICMS", text="Base de ICMS")
        tree_interestaduais.heading("Valor de ICMS", text="Valor de ICMS")

        for _, row in somas_interestaduais.iterrows():
            tree_interestaduais.insert("", tk.END, values=(row['Alíquota de ICMS'], row['Base de ICMS'], row['Valor do ICMS']))

        tree_interestaduais.pack(fill="both", expand=True, padx=10, pady=10)

    if not somas_estaduais.empty:
        # Frame para operações estaduais
        frame_estaduais = ttk.Frame(notebook)
        notebook.add(frame_estaduais, text="Operações Estaduais")

        tree_estaduais = ttk.Treeview(frame_estaduais, columns=("Alíquota", "Base de ICMS", "Valor de ICMS"), show="headings")
        tree_estaduais.heading("Alíquota", text="Alíquota de ICMS")
        tree_estaduais.heading("Base de ICMS", text="Base de ICMS")
        tree_estaduais.heading("Valor de ICMS", text="Valor de ICMS")

        for _, row in somas_estaduais.iterrows():
            tree_estaduais.insert("", tk.END, values=(row['Alíquota de ICMS'], row['Base de ICMS'], row['Valor do ICMS']))

        tree_estaduais.pack(fill="both", expand=True, padx=10, pady=10)

    # Frame para os botões de exportação
    frame_botoes = ttk.Frame(resultado_window)
    frame_botoes.pack(pady=10)

    if not somas_interestaduais.empty:
        btn_export_interestaduais = ttk.Button(frame_botoes, text="Exportar Interestaduais", command=lambda: exportar_dados(somas_interestaduais, "interestaduais"))
        btn_export_interestaduais.pack(side=tk.LEFT, padx=10)

    if not somas_estaduais.empty:
        btn_export_estaduais = ttk.Button(frame_botoes, text="Exportar Estaduais", command=lambda: exportar_dados(somas_estaduais, "estaduais"))
        btn_export_estaduais.pack(side=tk.LEFT, padx=10)

# Função para exportar os dados filtrados
def exportar_dados(df, tipo):
    save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")], title=f"Salvar Arquivo CSV - {tipo.capitalize()}")
    if save_path:
        try:
            df.to_csv(save_path, index=False, sep=';', encoding='latin1')
            messagebox.showinfo("Sucesso", f"Dados exportados com sucesso para {save_path}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar os dados: {e}")

# Função para sair do aplicativo
def sair():
    resposta = messagebox.askyesno("Sair", "Tem certeza que deseja sair?")
    if resposta:
        root.destroy()

# Configuração da interface gráfica
root = tk.Tk()
root.title("Análise de ICMS")

# Definir o tamanho mínimo da janela
root.minsize(400, 200)

# Adicionar um menu
menu_bar = tk.Menu(root)
root.config(menu=menu_bar)

# Menu Arquivo
arquivo_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Arquivo", menu=arquivo_menu)
arquivo_menu.add_command(label="Carregar Arquivo CSV", command=carregar_arquivo)
arquivo_menu.add_separator()
arquivo_menu.add_command(label="Sair", command=sair)

# Menu Ajuda
ajuda_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Ajuda", menu=ajuda_menu)
ajuda_menu.add_command(label="Sobre", command=lambda: messagebox.showinfo("Sobre", "Análise de ICMS\nVersão 1.0"))

# Frame principal
main_frame = ttk.Frame(root, padding="20")
main_frame.pack(fill='both', expand=True)

# Texto de boas-vindas
lbl_bem_vindo = ttk.Label(main_frame, text="Bem-vindo ao Analisador de ICMS", font=("Arial", 16))
lbl_bem_vindo.pack(pady=10)

# Botão para carregar o arquivo
btn_carregar = ttk.Button(main_frame, text="Carregar Arquivo CSV", command=carregar_arquivo)
btn_carregar.pack(pady=10)

# Iniciar a interface
root.mainloop()
