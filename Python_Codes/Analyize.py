import os
import re
import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
import threading

def selecionar_pasta_pdfs():
    pasta = filedialog.askdirectory(title="Selecione a pasta com os PDFs")
    if pasta:
        entrada_pasta_pdfs.delete(0, tk.END)
        entrada_pasta_pdfs.insert(0, pasta)

def selecionar_excel_empresas():
    arquivo = filedialog.askopenfilename(
        title="Selecione o Excel com as empresas filtradas",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    if arquivo:
        entrada_excel.delete(0, tk.END)
        entrada_excel.insert(0, arquivo)

def selecionar_pasta_saida():
    pasta = filedialog.askdirectory(title="Selecione a pasta para salvar o resultado")
    if pasta:
        entrada_pasta_saida.delete(0, tk.END)
        entrada_pasta_saida.insert(0, pasta)

def normalizar_cnpj(cnpj):
    """Remove formatação do CNPJ e retorna apenas números"""
    if not cnpj:
        return ""
    return re.sub(r'\D', '', cnpj)

def extrair_dados_pdf(caminho_pdf):
    dados = []
    nome_arquivo = os.path.basename(caminho_pdf)
    
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            texto = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())

            # Extrair CNPJ e Nome da empresa
            cnpj_match = re.search(r"CNPJ:\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto)
            cnpj_formatado = cnpj_match.group(1) if cnpj_match else "Não encontrado"
            cnpj_numeros = normalizar_cnpj(cnpj_formatado)
            
            # Extrair nome da empresa
            nome_match = re.search(r"CNPJ:\s*\d{2}\.\d{3}\.\d{3}.*?-\s*(.+)", texto)
            nome_empresa = nome_match.group(1).strip() if nome_match else "Não encontrado"

            # 1) PARCSN/PARCMEI - MEI
            if "MEI - EM PARCELAMENTO" in texto:
                mei_match = re.search(r"MEI - EM PARCELAMENTO\s+Parcelas em atraso\s*(\d+)", texto)
                if mei_match:
                    dados.append({
                        "CNPJ": cnpj_formatado,
                        "CNPJ_Numeros": cnpj_numeros,
                        "Nome_Empresa": nome_empresa,
                        "Tipo": "PARCMEI",
                        "Subtipo": "MEI",
                        "Conta": "-",
                        "Modalidade": "MEI - Parcelamento",
                        "Detalhes": f"Parcelas em atraso: {mei_match.group(1)}",
                        "Status": "Em Parcelamento",
                        "Arquivo": nome_arquivo
                    })

            # 2) PARCSN/PARCMEI - Simples Nacional
            if "SIMPLES NACIONAL - EM PARCELAMENTO" in texto:
                sn_match = re.search(r"SIMPLES NACIONAL - EM PARCELAMENTO\s+Parcelas em atraso\s*(\d+)", texto)
                if sn_match:
                    dados.append({
                        "CNPJ": cnpj_formatado,
                        "CNPJ_Numeros": cnpj_numeros,
                        "Nome_Empresa": nome_empresa,
                        "Tipo": "PARCSN",
                        "Subtipo": "Simples Nacional",
                        "Conta": "-",
                        "Modalidade": "Simples Nacional - Parcelamento",
                        "Detalhes": f"Parcelas em atraso: {sn_match.group(1)}",
                        "Status": "Em Parcelamento",
                        "Arquivo": nome_arquivo
                    })

            # 3) SIEFPAR - Parcelamento com Exigibilidade Suspensa (Receita Federal)
            if "Parcelamento com Exigibilidade Suspensa (SIEFPAR)" in texto:
                # Padrão melhorado para capturar parcelamentos
                siefpar_pattern = r"Parcelamento:\s*(\d+)\s+Valor Suspenso:\s*([\d\.,]+)\s*(.+?)(?=Parcelamento:|$)"
                matches = re.findall(siefpar_pattern, texto, re.DOTALL)
                
                for parcela, valor, detalhes in matches:
                    modalidade = "Parcelamento Simplificado"
                    if "Parcelamento Simplificado" in detalhes:
                        modalidade = "Parcelamento Simplificado"
                    
                    dados.append({
                        "CNPJ": cnpj_formatado,
                        "CNPJ_Numeros": cnpj_numeros,
                        "Nome_Empresa": nome_empresa,
                        "Tipo": "SIEFPAR",
                        "Subtipo": "Receita Federal",
                        "Conta": parcela.strip(),
                        "Modalidade": modalidade,
                        "Detalhes": f"Valor suspenso: R$ {valor}",
                        "Status": "Exigibilidade Suspensa",
                        "Arquivo": nome_arquivo
                    })

            # 4) SISPAR - Parcelamento com Exigibilidade Suspensa (PGFN)
            if "Parcelamento com Exigibilidade Suspensa (SISPAR)" in texto:
                # Busca por contas SISPAR
                sispar_pattern = r"Conta\s+(\d+)\s+(.+?)\s+Modalidade:\s*(.+?)(?=\n|$)"
                matches = re.findall(sispar_pattern, texto, re.MULTILINE)
                
                for conta, tipo_parcela, modalidade in matches:
                    dados.append({
                        "CNPJ": cnpj_formatado,
                        "CNPJ_Numeros": cnpj_numeros,
                        "Nome_Empresa": nome_empresa,
                        "Tipo": "SISPAR",
                        "Subtipo": "PGFN",
                        "Conta": conta.strip(),
                        "Modalidade": modalidade.strip(),
                        "Detalhes": tipo_parcela.strip(),
                        "Status": "Exigibilidade Suspensa",
                        "Arquivo": nome_arquivo
                    })

            # 5) SICOB - Débito com Exigibilidade Suspensa
            if "Débito com Exigibilidade Suspensa (SICOB)" in texto:
                sicob_pattern = r"Parcelamento:\s*(\d+-\d+)\s+Situação:\s*(\d+\s*-\s*.+)"
                matches = re.findall(sicob_pattern, texto)
                
                for parcela, situacao in matches:
                    dados.append({
                        "CNPJ": cnpj_formatado,
                        "CNPJ_Numeros": cnpj_numeros,
                        "Nome_Empresa": nome_empresa,
                        "Tipo": "SICOB",
                        "Subtipo": "Débito Suspenso",
                        "Conta": parcela.strip(),
                        "Modalidade": "RFB LEI 10522/02",
                        "Detalhes": f"Situação: {situacao}",
                        "Status": "Ativo/Em Dia",
                        "Arquivo": nome_arquivo
                    })

            # 6) Inscrições SIDA com parcelamento
            # if "NEGOCIADA NO SISPAR" in texto:
            #     sida_pattern = r"(\d{2}\.\d+\.\d{2}\.\d+-\d{2})\s+.+?\s+\d{2}/\d{2}/\d{4}.+?NEGOCIADA NO SISPAR"
            #     matches = re.findall(sida_pattern, texto, re.DOTALL)
                
            #     for inscricao in matches:
            #         dados.append({
            #             "CNPJ": cnpj_formatado,
            #             "CNPJ_Numeros": cnpj_numeros,
            #             "Nome_Empresa": nome_empresa,
            #             "Tipo": "SIDA",
            #             "Subtipo": "Inscrição Negociada",
            #             "Conta": inscricao,
            #             "Modalidade": "Negociada no SISPAR",
            #             "Detalhes": "Inscrição com exigibilidade suspensa",
            #             "Status": "Negociada",
            #             "Arquivo": nome_arquivo
            #         })

    except Exception as e:
        print(f"Erro ao processar {nome_arquivo}: {str(e)}")
        
    return dados

def processar_pdfs():
    pasta_pdfs = entrada_pasta_pdfs.get()
    excel_empresas = entrada_excel.get()
    pasta_saida = entrada_pasta_saida.get()

    if not pasta_pdfs:
        messagebox.showerror("Erro", "Selecione a pasta com os PDFs!")
        return

    if not pasta_saida:
        messagebox.showerror("Erro", "Selecione a pasta de saída!")
        return

    # Limpa a área de resultados
    text_resultados.delete(1.0, tk.END)
    
    # Limpa a tabela
    for item in tree_parcelamentos.get_children():
        tree_parcelamentos.delete(item)

    def processar():
        try:
            # Carrega lista de empresas se fornecida
            empresas_filtradas = set()
            if excel_empresas:
                try:
                    df_empresas = pd.read_excel(excel_empresas, dtype=str)
                    if 'CNPJ' in df_empresas.columns:
                        empresas_filtradas = set(df_empresas['CNPJ'].apply(normalizar_cnpj))
                    text_resultados.insert(tk.END, f"✅ Carregadas {len(empresas_filtradas)} empresas do Excel\n\n")
                except Exception as e:
                    text_resultados.insert(tk.END, f"⚠️ Erro ao carregar Excel: {str(e)}\n\n")

            todos_dados = []
            arquivos_pdf = [f for f in os.listdir(pasta_pdfs) if f.lower().endswith('.pdf')]
            
            text_resultados.insert(tk.END, f"📁 Encontrados {len(arquivos_pdf)} PDFs para processar...\n\n")
            janela.update()

            for i, arquivo in enumerate(arquivos_pdf, 1):
                caminho = os.path.join(pasta_pdfs, arquivo)
                text_resultados.insert(tk.END, f"[{i}/{len(arquivos_pdf)}] Processando: {arquivo}\n")
                janela.update()
                
                dados = extrair_dados_pdf(caminho)
                
                # Se há filtro de empresas, aplica
                if empresas_filtradas:
                    dados_filtrados = []
                    for d in dados:
                        if d['CNPJ_Numeros'] in empresas_filtradas:
                            dados_filtrados.append(d)
                            text_resultados.insert(tk.END, f"  ✅ Empresa encontrada na lista: {d['Nome_Empresa']}\n")
                    dados = dados_filtrados
                
                if dados:
                    text_resultados.insert(tk.END, f"  📊 {len(dados)} parcelamentos encontrados\n")
                else:
                    text_resultados.insert(tk.END, f"  ⚠️ Nenhum parcelamento encontrado\n")
                
                todos_dados.extend(dados)
                janela.update()

            # Cria DataFrame e salva
            if todos_dados:
                df = pd.DataFrame(todos_dados)
                
                # Ordena por empresa e tipo
                df = df.sort_values(['Nome_Empresa', 'Tipo'])
                
                # Salva Excel
                caminho_excel = os.path.join(pasta_saida, "parcelamentos_detalhados.xlsx")
                df.to_excel(caminho_excel, index=False)
                
                # Preenche a tabela na interface
                for _, row in df.iterrows():
                    tree_parcelamentos.insert("", tk.END, values=(
                        row['Nome_Empresa'],
                        row['CNPJ'],
                        row['Tipo'],
                        row['Subtipo'],
                        row['Conta'],
                        row['Modalidade'],
                        row['Status'],
                        row['Detalhes']
                    ))
                
                text_resultados.insert(tk.END, f"\n✅ CONCLUÍDO!\n")
                text_resultados.insert(tk.END, f"📊 Total de parcelamentos: {len(todos_dados)}\n")
                text_resultados.insert(tk.END, f"🏢 Empresas com parcelamentos: {df['Nome_Empresa'].nunique()}\n")
                text_resultados.insert(tk.END, f"💾 Arquivo salvo: {caminho_excel}\n")
                
                # Resumo por tipo
                text_resultados.insert(tk.END, f"\n📈 RESUMO POR TIPO:\n")
                resumo = df['Tipo'].value_counts()
                for tipo, qtd in resumo.items():
                    text_resultados.insert(tk.END, f"  • {tipo}: {qtd}\n")
                
            else:
                text_resultados.insert(tk.END, f"\n⚠️ Nenhum parcelamento foi encontrado nos PDFs!\n")

        except Exception as e:
            text_resultados.insert(tk.END, f"\n❌ ERRO: {str(e)}\n")
        
        finally:
            btn_processar.config(state="normal", text="🔄 Processar PDFs")

    # Executa em thread separada
    btn_processar.config(state="disabled", text="Processando...")
    thread = threading.Thread(target=processar)
    thread.daemon = True
    thread.start()

def filtrar_tabela():
    filtro = entrada_filtro.get().lower()
    
    # Limpa a tabela
    for item in tree_parcelamentos.get_children():
        tree_parcelamentos.delete(item)
    
    # Se não há filtro, não faz nada
    if not filtro:
        return
    
    # Recarrega dados filtrados (isso é simplificado, em produção manteria os dados em memória)
    pass

def exportar_selecionados():
    selecionados = tree_parcelamentos.selection()
    if not selecionados:
        messagebox.showwarning("Aviso", "Selecione pelo menos um item!")
        return
    
    # Implementar exportação dos itens selecionados
    messagebox.showinfo("Info", f"{len(selecionados)} itens selecionados para exportação")

# Criar janela principal
janela = tk.Tk()
janela.title("Analisador de Parcelamentos - PDFs da Receita Federal")
janela.geometry("1400x800")
janela.resizable(True, True)

# Notebook para abas
notebook = ttk.Notebook(janela)
notebook.pack(fill="both", expand=True, padx=10, pady=10)

# Aba 1: Configuração
frame_config = ttk.Frame(notebook)
notebook.add(frame_config, text="⚙️ Configuração")

# Título
title_label = tk.Label(frame_config, text="📊 Analisador de Parcelamentos", 
                      font=("Arial", 16, "bold"), fg="#2E7D32")
title_label.pack(pady=(10, 20))

# Pasta dos PDFs
tk.Label(frame_config, text="📁 Pasta com os PDFs:", font=("Arial", 10, "bold")).pack(anchor="w", padx=20, pady=(5, 2))
frame_pdfs = tk.Frame(frame_config)
frame_pdfs.pack(fill="x", padx=20, pady=(0, 10))
entrada_pasta_pdfs = tk.Entry(frame_pdfs, font=("Arial", 9))
entrada_pasta_pdfs.pack(side="left", fill="x", expand=True, padx=(0, 5))
tk.Button(frame_pdfs, text="Selecionar", command=selecionar_pasta_pdfs, 
          bg="#1976D2", fg="white").pack(side="right")

# Excel de empresas (opcional)
tk.Label(frame_config, text="📊 Excel com empresas filtradas (opcional):", font=("Arial", 10, "bold")).pack(anchor="w", padx=20, pady=(5, 2))
frame_excel = tk.Frame(frame_config)
frame_excel.pack(fill="x", padx=20, pady=(0, 10))
entrada_excel = tk.Entry(frame_excel, font=("Arial", 9))
entrada_excel.pack(side="left", fill="x", expand=True, padx=(0, 5))
tk.Button(frame_excel, text="Selecionar", command=selecionar_excel_empresas, 
          bg="#1976D2", fg="white").pack(side="right")

# Pasta de saída
tk.Label(frame_config, text="💾 Pasta para salvar resultado:", font=("Arial", 10, "bold")).pack(anchor="w", padx=20, pady=(5, 2))
frame_saida = tk.Frame(frame_config)
frame_saida.pack(fill="x", padx=20, pady=(0, 20))
entrada_pasta_saida = tk.Entry(frame_saida, font=("Arial", 9))
entrada_pasta_saida.pack(side="left", fill="x", expand=True, padx=(0, 5))
tk.Button(frame_saida, text="Selecionar", command=selecionar_pasta_saida, 
          bg="#1976D2", fg="white").pack(side="right")

# Botão processar
btn_processar = tk.Button(frame_config, text="🔄 Processar PDFs", command=processar_pdfs, 
                         bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), 
                         padx=30, pady=10)
btn_processar.pack(pady=20)

# Área de resultados do processamento
tk.Label(frame_config, text="📋 Log do Processamento:", font=("Arial", 10, "bold")).pack(anchor="w", padx=20, pady=(20, 5))
text_resultados = ScrolledText(frame_config, height=15, font=("Courier", 9))
text_resultados.pack(fill="both", expand=True, padx=20, pady=(0, 20))

# Aba 2: Resultados
frame_resultados = ttk.Frame(notebook)
notebook.add(frame_resultados, text="📊 Parcelamentos")

# Filtro
frame_filtro = tk.Frame(frame_resultados)
frame_filtro.pack(fill="x", padx=10, pady=10)
tk.Label(frame_filtro, text="🔍 Filtrar:", font=("Arial", 10)).pack(side="left", padx=(0, 5))
entrada_filtro = tk.Entry(frame_filtro, font=("Arial", 9))
entrada_filtro.pack(side="left", padx=(0, 10))
tk.Button(frame_filtro, text="Filtrar", command=filtrar_tabela, bg="#2196F3", fg="white").pack(side="left", padx=(0, 10))
tk.Button(frame_filtro, text="Exportar Selecionados", command=exportar_selecionados, bg="#FF9800", fg="white").pack(side="right")

# Tabela de parcelamentos
frame_tabela = tk.Frame(frame_resultados)
frame_tabela.pack(fill="both", expand=True, padx=10, pady=(0, 10))

colunas = ("Empresa", "CNPJ", "Tipo", "Subtipo", "Conta", "Modalidade", "Status", "Detalhes")
tree_parcelamentos = ttk.Treeview(frame_tabela, columns=colunas, show="headings", height=20)

# Configurar colunas
tree_parcelamentos.heading("Empresa", text="Empresa")
tree_parcelamentos.heading("CNPJ", text="CNPJ")
tree_parcelamentos.heading("Tipo", text="Tipo")
tree_parcelamentos.heading("Subtipo", text="Subtipo")
tree_parcelamentos.heading("Conta", text="Conta")
tree_parcelamentos.heading("Modalidade", text="Modalidade")
tree_parcelamentos.heading("Status", text="Status")
tree_parcelamentos.heading("Detalhes", text="Detalhes")

tree_parcelamentos.column("Empresa", width=200)
tree_parcelamentos.column("CNPJ", width=130)
tree_parcelamentos.column("Tipo", width=80)
tree_parcelamentos.column("Subtipo", width=100)
tree_parcelamentos.column("Conta", width=100)
tree_parcelamentos.column("Modalidade", width=200)
tree_parcelamentos.column("Status", width=120)
tree_parcelamentos.column("Detalhes", width=250)

# Scrollbars
scroll_y = ttk.Scrollbar(frame_tabela, orient="vertical", command=tree_parcelamentos.yview)
scroll_x = ttk.Scrollbar(frame_tabela, orient="horizontal", command=tree_parcelamentos.xview)
tree_parcelamentos.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

tree_parcelamentos.pack(side="left", fill="both", expand=True)
scroll_y.pack(side="right", fill="y")
scroll_x.pack(side="bottom", fill="x")

# Informações
info_text = """
ℹ️ TIPOS DE PARCELAMENTOS DETECTADOS:

• PARCMEI: Parcelamento MEI (Microempreendedor Individual)
• PARCSN: Parcelamento Simples Nacional
• SIEFPAR: Parcelamento com Exigibilidade Suspensa (Receita Federal)
• SISPAR: Parcelamento com Exigibilidade Suspensa (PGFN)
• SICOB: Débito com Exigibilidade Suspensa
• SIDA: Inscrições em Dívida Ativa negociadas

📋 Como usar:
1. Selecione a pasta contendo os PDFs da Receita Federal
2. (Opcional) Selecione o Excel com empresas filtradas
3. Escolha onde salvar o resultado
4. Clique em "Processar PDFs"
5. Acompanhe o progresso na aba "Configuração"
6. Visualize os resultados na aba "Parcelamentos"
"""

info_label = tk.Label(frame_config, text=info_text, font=("Arial", 8), 
                     fg="#555555", justify="left", anchor="w")
info_label.pack(fill="x", padx=20, pady=(10, 0))

janela.mainloop()