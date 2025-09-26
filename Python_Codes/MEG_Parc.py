import os
import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def selecionar_excel_dados_desejados():
    arquivo = filedialog.askopenfilename(
        title="Selecione o Excel com os dados que você QUER",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    if arquivo:
        entrada_excel_desejados.delete(0, tk.END)
        entrada_excel_desejados.insert(0, arquivo)

def selecionar_excel_todas_empresas():
    arquivo = filedialog.askopenfilename(
        title="Selecione o Excel com TODAS as empresas",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    if arquivo:
        entrada_excel_todas.delete(0, tk.END)
        entrada_excel_todas.insert(0, arquivo)

def selecionar_pasta_saida():
    pasta = filedialog.askdirectory(title="Selecione a pasta para salvar o resultado")
    if pasta:
        entrada_pasta_saida.delete(0, tk.END)
        entrada_pasta_saida.insert(0, pasta)

def normalize_cnpj(value):
    """
    Remove tudo que não for dígito e completa com zeros à esquerda até 14 dígitos.
    Se value for NaN ou vazio, retorna string vazia.
    """
    if pd.isna(value):
        return ""
    s = str(value)
    # Remove tudo que não for dígito
    digits = re.sub(r'\D', '', s)
    if digits == "":
        return ""
    # Se tiver menos de 14 dígitos, acrescenta zeros à esquerda até 14
    if len(digits) < 14:
        digits = digits.zfill(14)
    return digits

def normalize_nome(value):
    """
    Normaliza o nome removendo espaços extras, acentos e convertendo para minúsculas.
    """
    if pd.isna(value):
        return ""
    nome = str(value).strip().lower()
    # Remove acentos e caracteres especiais básicos
    nome = (nome.replace('ã', 'a').replace('á', 'a').replace('à', 'a').replace('â', 'a')
                .replace('é', 'e').replace('ê', 'e').replace('í', 'i').replace('ó', 'o')
                .replace('ô', 'o').replace('õ', 'o').replace('ú', 'u').replace('ü', 'u')
                .replace('ç', 'c').replace('ñ', 'n'))
    # Remove espaços múltiplos
    nome = re.sub(r'\s+', ' ', nome)
    return nome

def gerar_excel():
    excel_desejados = entrada_excel_desejados.get()
    excel_todas = entrada_excel_todas.get()
    pasta_saida = entrada_pasta_saida.get()

    if not excel_desejados or not excel_todas or not pasta_saida:
        messagebox.showerror("Erro", "Selecione todos os campos!")
        return

    try:
        # Atualiza o status
        status_label.config(text="Carregando Excel com dados desejados...")
        janela.update()

        # Lê o Excel com dados desejados (coluna A = Nome, coluna B = CNPJ)
        df_desejados = pd.read_excel(excel_desejados, dtype=str, header=None)
        
        # Verifica se tem pelo menos 2 colunas
        if len(df_desejados.columns) < 2:
            messagebox.showerror("Erro", "O Excel de dados desejados deve ter pelo menos 2 colunas (A=Nome, B=CNPJ)!")
            return

        # Define colunas A e B
        df_desejados.columns = [f'Col_{i}' for i in range(len(df_desejados.columns))]
        df_desejados['Nome_Desejado'] = df_desejados['Col_0']  # Coluna A
        df_desejados['CNPJ_Desejado'] = df_desejados['Col_1']  # Coluna B

        # Normaliza dados desejados
        df_desejados['Nome_Normalizado'] = df_desejados['Nome_Desejado'].apply(normalize_nome)
        df_desejados['CNPJ_Normalizado'] = df_desejados['CNPJ_Desejado'].apply(normalize_cnpj)

        # Remove linhas com nome vazio
        df_desejados = df_desejados[df_desejados['Nome_Normalizado'] != '']

        status_label.config(text="Carregando Excel com todas as empresas...")
        janela.update()

        # Lê o Excel com todas as empresas (coluna A = Código, coluna B = Nome, coluna C = CNPJ)
        df_todas = pd.read_excel(excel_todas, dtype=str, header=None)
        
        # Verifica se tem pelo menos 3 colunas
        if len(df_todas.columns) < 3:
            messagebox.showerror("Erro", "O Excel de todas as empresas deve ter pelo menos 3 colunas (A=Código, B=Nome, C=CNPJ)!")
            return

        # Define colunas A, B e C
        df_todas.columns = [f'Col_{i}' for i in range(len(df_todas.columns))]
        df_todas['Codigo'] = df_todas['Col_0']      # Coluna A
        df_todas['Nome_Todas'] = df_todas['Col_1']   # Coluna B
        df_todas['CNPJ_Todas'] = df_todas['Col_2']   # Coluna C

        # Normaliza dados de todas as empresas
        df_todas['Nome_Todas_Normalizado'] = df_todas['Nome_Todas'].apply(normalize_nome)

        # Remove linhas com nome vazio
        df_todas = df_todas[df_todas['Nome_Todas_Normalizado'] != '']

        status_label.config(text="Comparando dados por nome...")
        janela.update()

        # Cria conjunto de nomes desejados para busca
        nomes_desejados = set(df_desejados['Nome_Normalizado'])

        # Encontra empresas que batem pelo nome
        df_todas['Match'] = df_todas['Nome_Todas_Normalizado'].isin(nomes_desejados)
        df_encontradas = df_todas[df_todas['Match']].copy()

        if df_encontradas.empty:
            messagebox.showwarning("Aviso", "Nenhuma empresa foi encontrada com os nomes fornecidos!")
            return

        status_label.config(text="Montando resultado final...")
        janela.update()

        # Para cada empresa encontrada, busca os dados correspondentes no Excel desejado
        resultado_final = []

        for _, empresa in df_encontradas.iterrows():
            nome_normalizado = empresa['Nome_Todas_Normalizado']
            
            # Busca dados correspondentes no Excel desejado
            dados_desejados = df_desejados[df_desejados['Nome_Normalizado'] == nome_normalizado]
            
            if not dados_desejados.empty:
                # Pega o primeiro match (caso haja duplicatas)
                dados = dados_desejados.iloc[0]
                
                resultado_final.append({
                    'Nº': empresa['Codigo'],
                    'EMPRESA': dados['Nome_Desejado'],
                    'CNPJ': dados['CNPJ_Normalizado']
                })

        if not resultado_final:
            messagebox.showwarning("Aviso", "Nenhuma correspondência foi encontrada!")
            return

        # Cria DataFrame final
        df_resultado = pd.DataFrame(resultado_final)

        # Remove linhas com código vazio
        df_resultado = df_resultado[df_resultado['Nº'].astype(str).str.strip() != '']

        # Ordena por código
        df_resultado = df_resultado.sort_values('Nº').reset_index(drop=True)

        status_label.config(text="Salvando resultado...")
        janela.update()

        # Salva o resultado
        caminho_saida = os.path.join(pasta_saida, "empresas_filtradas.xlsx")
        df_resultado.to_excel(caminho_saida, index=False)

        # Limpa a tabela
        for item in tabela.get_children():
            tabela.delete(item)

        # Preenche a tabela na interface
        for _, row in df_resultado.iterrows():
            tabela.insert("", tk.END, values=(row['Nº'], row['EMPRESA'], row['CNPJ']))

        status_label.config(text=f"✅ Concluído! {len(df_resultado)} empresas processadas")
        
        messagebox.showinfo("Sucesso", 
                          f"Novo Excel gerado com sucesso!\n"
                          f"📁 Local: {caminho_saida}\n"
                          f"📊 Total de empresas: {len(df_resultado)}\n\n"
                          f"Estrutura do resultado:\n"
                          f"• Código: do Excel 2 (todas as empresas)\n"
                          f"• Nome e CNPJ: do Excel 1 (dados desejados)")

    except Exception as e:
        status_label.config(text="❌ Erro durante o processamento")
        messagebox.showerror("Erro", f"Ocorreu um erro:\n{str(e)}")

def limpar_campos():
    entrada_excel_desejados.delete(0, tk.END)
    entrada_excel_todas.delete(0, tk.END)
    entrada_pasta_saida.delete(0, tk.END)
    for item in tabela.get_children():
        tabela.delete(item)
    status_label.config(text="Campos limpos")

# Criar janela
janela = tk.Tk()
janela.title("Filtro de Empresas - Comparação por Colunas Específicas")
janela.geometry("850x650")
janela.resizable(True, True)

# Frame principal
main_frame = tk.Frame(janela)
main_frame.pack(fill="both", expand=True, padx=10, pady=10)

# Título
title_label = tk.Label(main_frame, text="📊 Filtro de Empresas por Colunas A e B", 
                      font=("Arial", 16, "bold"), fg="#2E7D32")
title_label.pack(pady=(0, 20))

# Seção Excel com dados desejados
tk.Label(main_frame, text="1️⃣ Excel com dados que você QUER (A=Nome, B=CNPJ):", 
         font=("Arial", 10, "bold")).pack(anchor="w", pady=(5, 2))
frame_desejados = tk.Frame(main_frame)
frame_desejados.pack(fill="x", pady=(0, 10))
entrada_excel_desejados = tk.Entry(frame_desejados, font=("Arial", 9))
entrada_excel_desejados.pack(side="left", fill="x", expand=True, padx=(0, 5))
tk.Button(frame_desejados, text="Selecionar", command=selecionar_excel_dados_desejados, 
          bg="#1976D2", fg="white").pack(side="right")

# Seção Excel com todas as empresas
tk.Label(main_frame, text="2️⃣ Excel com TODAS as empresas (A=Código, B=Nome, C=CNPJ):", 
         font=("Arial", 10, "bold")).pack(anchor="w", pady=(5, 2))
frame_todas = tk.Frame(main_frame)
frame_todas.pack(fill="x", pady=(0, 10))
entrada_excel_todas = tk.Entry(frame_todas, font=("Arial", 9))
entrada_excel_todas.pack(side="left", fill="x", expand=True, padx=(0, 5))
tk.Button(frame_todas, text="Selecionar", command=selecionar_excel_todas_empresas, 
          bg="#1976D2", fg="white").pack(side="right")

# Seção pasta de saída
tk.Label(main_frame, text="3️⃣ Pasta para salvar o resultado:", 
         font=("Arial", 10, "bold")).pack(anchor="w", pady=(5, 2))
frame_saida = tk.Frame(main_frame)
frame_saida.pack(fill="x", pady=(0, 15))
entrada_pasta_saida = tk.Entry(frame_saida, font=("Arial", 9))
entrada_pasta_saida.pack(side="left", fill="x", expand=True, padx=(0, 5))
tk.Button(frame_saida, text="Selecionar", command=selecionar_pasta_saida, 
          bg="#1976D2", fg="white").pack(side="right")

# Frame dos botões
frame_botoes = tk.Frame(main_frame)
frame_botoes.pack(pady=10)

tk.Button(frame_botoes, text="🔄 Processar", command=gerar_excel, 
          bg="#4CAF50", fg="white", font=("Arial", 10, "bold"), 
          padx=20, pady=5).pack(side="left", padx=(0, 10))

tk.Button(frame_botoes, text="🗑️ Limpar", command=limpar_campos, 
          bg="#FF9800", fg="white", font=("Arial", 10), 
          padx=20, pady=5).pack(side="left")

# Status
status_label = tk.Label(main_frame, text="Aguardando seleção dos arquivos...", 
                       font=("Arial", 9), fg="#666666")
status_label.pack(pady=(10, 5))

# Frame da tabela
frame_tabela = tk.Frame(main_frame)
frame_tabela.pack(fill="both", expand=True, pady=(10, 0))

# Tabela de resultados
tk.Label(frame_tabela, text="📋 Preview do resultado:", 
         font=("Arial", 10, "bold")).pack(anchor="w")

colunas = ("Nº", "EMPRESA", "CNPJ")
tabela = ttk.Treeview(frame_tabela, columns=colunas, show="headings", height=10)
for col in colunas:
    tabela.heading(col, text=col)
    if col == "Nº":
        tabela.column(col, width=80)
    elif col == "EMPRESA":
        tabela.column(col, width=300)
    else:  # CNPJ
        tabela.column(col, width=150)

# Scrollbar para a tabela
scrollbar = ttk.Scrollbar(frame_tabela, orient="vertical", command=tabela.yview)
tabela.configure(yscrollcommand=scrollbar.set)

tabela.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

# Informações de uso
info_text = """
ℹ️ Como usar:
1. Selecione o Excel com dados que você QUER (coluna A=Nome, coluna B=CNPJ)
2. Selecione o Excel com TODAS as empresas (coluna A=Código, coluna B=Nome, coluna C=CNPJ)
3. Escolha onde salvar o resultado
4. Clique em "Processar"

O sistema irá:
• Comparar os nomes da coluna B (Excel 2) com coluna A (Excel 1)
• Para cada match encontrado, pegar:
  - Código: da coluna A do Excel 2
  - Nome e CNPJ: das colunas A e B do Excel 1
• Normalizar os CNPJs (apenas números, completando com zeros)
• Gerar um novo Excel com: Nº, EMPRESA, CNPJ
"""

info_label = tk.Label(main_frame, text=info_text, font=("Arial", 8), 
                     fg="#555555", justify="left", anchor="w")
info_label.pack(pady=(10, 0), fill="x")

janela.mainloop()