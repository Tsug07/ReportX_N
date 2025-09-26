import os
import re
import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
import threading
from datetime import datetime
import json

class AnalisadorParcelamentos:
    def __init__(self):
        self.dados_processados = []
        self.empresas_filtradas = set()
        self.setup_gui()
        
    def setup_gui(self):
        # Criar janela principal
        self.janela = tk.Tk()
        self.janela.title("Analisador de Parcelamentos - PDFs da Receita Federal v2.0")
        self.janela.geometry("1500x900")
        self.janela.resizable(True, True)

        # Notebook para abas
        self.notebook = ttk.Notebook(self.janela)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.criar_aba_configuracao()
        self.criar_aba_resultados()
        self.criar_aba_dashboard()
        
    def criar_aba_configuracao(self):
        # Aba 1: Configuração
        frame_config = ttk.Frame(self.notebook)
        self.notebook.add(frame_config, text="⚙️ Configuração")

        # Título
        title_label = tk.Label(frame_config, text="📊 Analisador de Parcelamentos v2.0", 
                              font=("Arial", 16, "bold"), fg="#2E7D32")
        title_label.pack(pady=(10, 20))

        # Frame para inputs
        inputs_frame = tk.Frame(frame_config)
        inputs_frame.pack(fill="x", padx=20, pady=10)

        # Pasta dos PDFs
        tk.Label(inputs_frame, text="📁 Pasta com os PDFs:", font=("Arial", 10, "bold")).pack(anchor="w", pady=(5, 2))
        frame_pdfs = tk.Frame(inputs_frame)
        frame_pdfs.pack(fill="x", pady=(0, 10))
        self.entrada_pasta_pdfs = tk.Entry(frame_pdfs, font=("Arial", 9))
        self.entrada_pasta_pdfs.pack(side="left", fill="x", expand=True, padx=(0, 5))
        tk.Button(frame_pdfs, text="Selecionar", command=self.selecionar_pasta_pdfs, 
                  bg="#1976D2", fg="white").pack(side="right")

        # Excel de empresas (opcional)
        tk.Label(inputs_frame, text="📊 Excel com empresas filtradas (opcional):", font=("Arial", 10, "bold")).pack(anchor="w", pady=(5, 2))
        frame_excel = tk.Frame(inputs_frame)
        frame_excel.pack(fill="x", pady=(0, 10))
        self.entrada_excel = tk.Entry(frame_excel, font=("Arial", 9))
        self.entrada_excel.pack(side="left", fill="x", expand=True, padx=(0, 5))
        tk.Button(frame_excel, text="Selecionar", command=self.selecionar_excel_empresas, 
                  bg="#1976D2", fg="white").pack(side="right")

        # Pasta de saída
        tk.Label(inputs_frame, text="💾 Pasta para salvar resultado:", font=("Arial", 10, "bold")).pack(anchor="w", pady=(5, 2))
        frame_saida = tk.Frame(inputs_frame)
        frame_saida.pack(fill="x", pady=(0, 10))
        self.entrada_pasta_saida = tk.Entry(frame_saida, font=("Arial", 9))
        self.entrada_pasta_saida.pack(side="left", fill="x", expand=True, padx=(0, 5))
        tk.Button(frame_saida, text="Selecionar", command=self.selecionar_pasta_saida, 
                  bg="#1976D2", fg="white").pack(side="right")

        # Opções avançadas
        options_frame = tk.LabelFrame(inputs_frame, text="Opções Avançadas", font=("Arial", 9, "bold"))
        options_frame.pack(fill="x", pady=10)

        self.incluir_detalhes_debitos = tk.BooleanVar(value=True)
        tk.Checkbutton(options_frame, text="Incluir detalhes de débitos pendentes", 
                      variable=self.incluir_detalhes_debitos).pack(anchor="w", padx=10, pady=2)

        self.agrupar_por_empresa = tk.BooleanVar(value=False)
        tk.Checkbutton(options_frame, text="Agrupar resultados por empresa", 
                      variable=self.agrupar_por_empresa).pack(anchor="w", padx=10, pady=2)

        self.salvar_backup_json = tk.BooleanVar(value=True)
        tk.Checkbutton(options_frame, text="Salvar backup em JSON", 
                      variable=self.salvar_backup_json).pack(anchor="w", padx=10, pady=2)

        # Botões de ação
        buttons_frame = tk.Frame(inputs_frame)
        buttons_frame.pack(pady=20)

        self.btn_processar = tk.Button(buttons_frame, text="🔄 Processar PDFs", command=self.processar_pdfs, 
                                      bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), 
                                      padx=30, pady=10)
        self.btn_processar.pack(side="left", padx=10)

        tk.Button(buttons_frame, text="🧹 Limpar Resultados", command=self.limpar_resultados, 
                  bg="#FF5722", fg="white", font=("Arial", 10), padx=20, pady=10).pack(side="left", padx=10)

        tk.Button(buttons_frame, text="💾 Salvar Configuração", command=self.salvar_config, 
                  bg="#2196F3", fg="white", font=("Arial", 10), padx=20, pady=10).pack(side="left", padx=10)

        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(inputs_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x", pady=10)

        # Status atual
        self.status_label = tk.Label(inputs_frame, text="Aguardando configuração...", 
                                    font=("Arial", 9), fg="#666666")
        self.status_label.pack(pady=5)

        # Área de resultados do processamento
        tk.Label(inputs_frame, text="📋 Log do Processamento:", font=("Arial", 10, "bold")).pack(anchor="w", pady=(20, 5))
        self.text_resultados = ScrolledText(inputs_frame, height=12, font=("Courier", 9))
        self.text_resultados.pack(fill="both", expand=True, pady=(0, 20))

    def criar_aba_resultados(self):
        # Aba 2: Resultados
        frame_resultados = ttk.Frame(self.notebook)
        self.notebook.add(frame_resultados, text="📊 Parcelamentos")

        # Toolbar
        toolbar = tk.Frame(frame_resultados)
        toolbar.pack(fill="x", padx=10, pady=5)

        # Filtros
        tk.Label(toolbar, text="🔍 Filtros:", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=5, sticky="w")

        tk.Label(toolbar, text="Empresa:").grid(row=0, column=1, padx=5)
        self.filtro_empresa = tk.Entry(toolbar, width=20)
        self.filtro_empresa.grid(row=0, column=2, padx=5)

        tk.Label(toolbar, text="Tipo:").grid(row=0, column=3, padx=5)
        self.filtro_tipo = ttk.Combobox(toolbar, width=15, values=["Todos", "PARCMEI", "PARCSN", "SIEFPAR", "SISPAR", "SICOB"])
        self.filtro_tipo.set("Todos")
        self.filtro_tipo.grid(row=0, column=4, padx=5)

        tk.Button(toolbar, text="Aplicar Filtros", command=self.aplicar_filtros, bg="#2196F3", fg="white").grid(row=0, column=5, padx=10)
        tk.Button(toolbar, text="Limpar Filtros", command=self.limpar_filtros, bg="#FF9800", fg="white").grid(row=0, column=6, padx=5)

        # Ações
        tk.Button(toolbar, text="📤 Exportar Filtrados", command=self.exportar_filtrados, bg="#4CAF50", fg="white").grid(row=0, column=7, padx=10)
        tk.Button(toolbar, text="📋 Copiar Selecionados", command=self.copiar_selecionados, bg="#9C27B0", fg="white").grid(row=0, column=8, padx=5)

        # Contador de resultados
        self.label_contador = tk.Label(toolbar, text="Nenhum resultado", font=("Arial", 9), fg="#666")
        self.label_contador.grid(row=0, column=9, padx=20)

        # Tabela de parcelamentos
        frame_tabela = tk.Frame(frame_resultados)
        frame_tabela.pack(fill="both", expand=True, padx=10, pady=10)

        # Colunas da tabela
        colunas = ("Empresa", "CNPJ", "Tipo", "Subtipo", "Conta", "Modalidade", "Status", "Detalhes", "Valor", "Arquivo")
        self.tree_parcelamentos = ttk.Treeview(frame_tabela, columns=colunas, show="headings", height=20)

        # Configurar colunas
        larguras = {"Empresa": 180, "CNPJ": 120, "Tipo": 80, "Subtipo": 100, "Conta": 100, 
                   "Modalidade": 180, "Status": 120, "Detalhes": 200, "Valor": 100, "Arquivo": 150}
        
        for col in colunas:
            self.tree_parcelamentos.heading(col, text=col, command=lambda c=col: self.ordenar_coluna(c))
            self.tree_parcelamentos.column(col, width=larguras.get(col, 100))

        # Scrollbars
        scroll_y = ttk.Scrollbar(frame_tabela, orient="vertical", command=self.tree_parcelamentos.yview)
        scroll_x = ttk.Scrollbar(frame_tabela, orient="horizontal", command=self.tree_parcelamentos.xview)
        self.tree_parcelamentos.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self.tree_parcelamentos.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")

        # Context menu
        self.menu_contexto = tk.Menu(self.janela, tearoff=0)
        self.menu_contexto.add_command(label="Copiar CNPJ", command=self.copiar_cnpj)
        self.menu_contexto.add_command(label="Copiar linha completa", command=self.copiar_linha)
        self.menu_contexto.add_separator()
        self.menu_contexto.add_command(label="Abrir PDF", command=self.abrir_pdf)

        self.tree_parcelamentos.bind("<Button-3>", self.mostrar_menu_contexto)

    def criar_aba_dashboard(self):
        # Aba 3: Dashboard
        frame_dashboard = ttk.Frame(self.notebook)
        self.notebook.add(frame_dashboard, text="📈 Dashboard")

        # Estatísticas gerais
        stats_frame = tk.LabelFrame(frame_dashboard, text="Estatísticas Gerais", font=("Arial", 10, "bold"))
        stats_frame.pack(fill="x", padx=20, pady=10)

        # Grid para estatísticas
        self.stats_labels = {}
        stats_info = [
            ("total_empresas", "🏢 Total de Empresas:"),
            ("total_parcelamentos", "📊 Total de Parcelamentos:"),
            ("empresas_com_parcelas_atraso", "⚠️ Com Parcelas em Atraso:"),
            ("valor_total_suspenso", "💰 Valor Total Suspenso:")
        ]

        for i, (key, label) in enumerate(stats_info):
            tk.Label(stats_frame, text=label, font=("Arial", 9)).grid(row=i//2, column=(i%2)*2, sticky="w", padx=10, pady=5)
            self.stats_labels[key] = tk.Label(stats_frame, text="0", font=("Arial", 9, "bold"), fg="#2E7D32")
            self.stats_labels[key].grid(row=i//2, column=(i%2)*2+1, sticky="w", padx=10, pady=5)

        # Resumo por tipo
        resumo_frame = tk.LabelFrame(frame_dashboard, text="Resumo por Tipo de Parcelamento", font=("Arial", 10, "bold"))
        resumo_frame.pack(fill="both", expand=True, padx=20, pady=10)

        # Tabela resumo
        colunas_resumo = ("Tipo", "Quantidade", "Empresas", "Percentual")
        self.tree_resumo = ttk.Treeview(resumo_frame, columns=colunas_resumo, show="headings", height=8)
        
        for col in colunas_resumo:
            self.tree_resumo.heading(col, text=col)
            self.tree_resumo.column(col, width=120)

        self.tree_resumo.pack(fill="both", expand=True, padx=10, pady=10)

    # Métodos de interface
    def selecionar_pasta_pdfs(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta com os PDFs")
        if pasta:
            self.entrada_pasta_pdfs.delete(0, tk.END)
            self.entrada_pasta_pdfs.insert(0, pasta)

    def selecionar_excel_empresas(self):
        arquivo = filedialog.askopenfilename(
            title="Selecione o Excel com as empresas filtradas",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        if arquivo:
            self.entrada_excel.delete(0, tk.END)
            self.entrada_excel.insert(0, arquivo)

    def selecionar_pasta_saida(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta para salvar o resultado")
        if pasta:
            self.entrada_pasta_saida.delete(0, tk.END)
            self.entrada_pasta_saida.insert(0, pasta)

    def normalizar_cnpj(self, cnpj):
        """Remove formatação do CNPJ e retorna apenas números"""
        if not cnpj:
            return ""
        return re.sub(r'\D', '', cnpj)

    def extrair_valor_monetario(self, texto):
        """Extrai valores monetários do texto"""
        pattern = r'[\d\.,]+(?=\s*(?:reais?|R\$|\b))'
        matches = re.findall(pattern, texto)
        if matches:
            # Pega o maior valor encontrado
            valores = []
            for match in matches:
                try:
                    valor = float(match.replace('.', '').replace(',', '.'))
                    valores.append(valor)
                except:
                    continue
            return max(valores) if valores else 0
        return 0

    def extrair_dados_pdf(self, caminho_pdf):
        dados = []
        nome_arquivo = os.path.basename(caminho_pdf)
        
        try:
            with pdfplumber.open(caminho_pdf) as pdf:
                texto = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())

                # Extrair CNPJ e Nome da empresa
                cnpj_match = re.search(r"CNPJ:\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto)
                cnpj_formatado = cnpj_match.group(1) if cnpj_match else "Não encontrado"
                cnpj_numeros = self.normalizar_cnpj(cnpj_formatado)
                
                # Extrair nome da empresa
                nome_match = re.search(r"CNPJ:\s*\d{2}\.\d{3}\.\d{3}.*?-\s*(.+)", texto)
                nome_empresa = nome_match.group(1).strip() if nome_match else "Não encontrado"

                # 1) PARCMEI - MEI
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
                            "Valor": 0,
                            "Arquivo": nome_arquivo
                        })

                # 2) PARCSN - Simples Nacional
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
                            "Valor": 0,
                            "Arquivo": nome_arquivo
                        })

                # 3) SIEFPAR - Parcelamento com Exigibilidade Suspensa (Receita Federal)
                if "Parcelamento com Exigibilidade Suspensa (SIEFPAR)" in texto:
                    siefpar_pattern = r"Parcelamento:\s*(\d+)\s+Valor Suspenso:\s*([\d\.,]+)\s*(.+?)(?=Parcelamento:|$)"
                    matches = re.findall(siefpar_pattern, texto, re.DOTALL)
                    
                    for parcela, valor_str, detalhes in matches:
                        valor = self.extrair_valor_monetario(valor_str) if valor_str else 0
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
                            "Detalhes": f"Valor suspenso: R$ {valor_str}",
                            "Status": "Exigibilidade Suspensa",
                            "Valor": valor,
                            "Arquivo": nome_arquivo
                        })

                # 4) SISPAR - Parcelamento com Exigibilidade Suspensa (PGFN)
                if "Parcelamento com Exigibilidade Suspensa (SISPAR)" in texto:
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
                            "Valor": 0,
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
                            "Valor": 0,
                            "Arquivo": nome_arquivo
                        })

                # Incluir detalhes de débitos se solicitado
                if self.incluir_detalhes_debitos.get() and "Pendência - Débito (SIEF)" in texto:
                    # Extrair débitos pendentes
                    debito_pattern = r"(\d{4}-\d{2}\s*-\s*.+?)\s+(\d{2}/\d{4})\s+[\d/]+\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+(.+)"
                    debitos = re.findall(debito_pattern, texto)
                    
                    for receita, periodo, dt_vcto, vl_orig, sdo_dev, multa, juros, sdo_cons, situacao in debitos[:5]:  # Limita a 5 débitos
                        valor_total = self.extrair_valor_monetario(sdo_cons) if sdo_cons else 0
                        dados.append({
                            "CNPJ": cnpj_formatado,
                            "CNPJ_Numeros": cnpj_numeros,
                            "Nome_Empresa": nome_empresa,
                            "Tipo": "DÉBITO",
                            "Subtipo": "Pendência",
                            "Conta": receita.strip(),
                            "Modalidade": f"Período: {periodo}",
                            "Detalhes": f"Situação: {situacao.strip()}",
                            "Status": "Devedor",
                            "Valor": valor_total,
                            "Arquivo": nome_arquivo
                        })

        except Exception as e:
            self.log(f"Erro ao processar {nome_arquivo}: {str(e)}")
            
        return dados

    def processar_pdfs(self):
        pasta_pdfs = self.entrada_pasta_pdfs.get()
        excel_empresas = self.entrada_excel.get()
        pasta_saida = self.entrada_pasta_saida.get()

        if not pasta_pdfs:
            messagebox.showerror("Erro", "Selecione a pasta com os PDFs!")
            return

        if not pasta_saida:
            messagebox.showerror("Erro", "Selecione a pasta de saída!")
            return

        def processar():
            try:
                self.limpar_resultados()
                inicio = datetime.now()
                
                # Carrega lista de empresas se fornecida
                if excel_empresas:
                    try:
                        df_empresas = pd.read_excel(excel_empresas, dtype=str)
                        if 'CNPJ' in df_empresas.columns:
                            self.empresas_filtradas = set(df_empresas['CNPJ'].apply(self.normalizar_cnpj))
                        self.log(f"✅ Carregadas {len(self.empresas_filtradas)} empresas do Excel")
                    except Exception as e:
                        self.log(f"⚠️ Erro ao carregar Excel: {str(e)}")

                # Lista arquivos PDF
                arquivos_pdf = [f for f in os.listdir(pasta_pdfs) if f.lower().endswith('.pdf')]
                total_arquivos = len(arquivos_pdf)
                
                self.log(f"📁 Encontrados {total_arquivos} PDFs para processar...")
                self.progress_var.set(0)

                todos_dados = []
                for i, arquivo in enumerate(arquivos_pdf, 1):
                    caminho = os.path.join(pasta_pdfs, arquivo)
                    self.status_label.config(text=f"Processando {i}/{total_arquivos}: {arquivo}")
                    self.log(f"[{i}/{total_arquivos}] {arquivo}")
                    
                    dados = self.extrair_dados_pdf(caminho)
                    
                    # Aplicar filtro de empresas se existe
                    if self.empresas_filtradas:
                        dados_filtrados = [d for d in dados if d['CNPJ_Numeros'] in self.empresas_filtradas]
                        if dados_filtrados:
                            self.log(f"  ✅ {len(dados_filtrados)} parcelamentos encontrados (filtrado)")
                        dados = dados_filtrados
                    else:
                        if dados:
                            self.log(f"  📊 {len(dados)} parcelamentos encontrados")
                    
                    todos_dados.extend(dados)
                    
                    # Atualizar progress bar
                    progresso = (i / total_arquivos) * 100
                    self.progress_var.set(progresso)
                    self.janela.update()

                # Processar e salvar resultados
                if todos_dados:
                    self.dados_processados = todos_dados
                    df = pd.DataFrame(todos_dados)
                    
                    # Agrupar por empresa se solicitado
                    if self.agrupar_por_empresa.get():
                        df = df.sort_values(['Nome_Empresa', 'Tipo'])
                    else:
                        df = df.sort_values(['Tipo', 'Nome_Empresa'])
                    
                    # Salvar Excel
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    caminho_excel = os.path.join(pasta_saida, f"parcelamentos_detalhados_{timestamp}.xlsx")
                    df.to_excel(caminho_excel, index=False)
                    
                    # Salvar backup JSON se solicitado
                    if self.salvar_backup_json.get():
                        caminho_json = os.path.join(pasta_saida, f"parcelamentos_backup_{timestamp}.json")
                        with open(caminho_json, 'w', encoding='utf-8') as f:
                            json.dump(todos_dados, f, ensure_ascii=False, indent=2)
                    
                    # Atualizar interface
                    self.atualizar_tabela()
                    self.atualizar_dashboard()
                    
                    fim = datetime.now()
                    tempo_total = (fim - inicio).total_seconds()
                    
                    self.log(f"\n✅ PROCESSAMENTO CONCLUÍDO!")
                    self.log(f"⏱️ Tempo total: {tempo_total:.1f} segundos")
                    self.log(f"📊 Total de parcelamentos: {len(todos_dados)}")
                    self.log(f"🏢 Empresas processadas: {df['Nome_Empresa'].nunique()}")
                    self.log(f"💾 Arquivo salvo: {caminho_excel}")
                    
                else:
                    self.log(f"\n⚠️ Nenhum parcelamento foi encontrado nos PDFs!")

            except Exception as e:
                self.log(f"\n❌ ERRO: {str(e)}")
            
            finally:
                self.btn_processar.config(state="normal", text="🔄 Processar PDFs")
                self.status_label.config(text="Processamento concluído")
                self.progress_var.set(100)

        # Executar em thread separada
        self.btn_processar.config(state="disabled", text="Processando...")
        thread = threading.Thread(target=processar)
        thread.daemon = True
        thread.start()

    def log(self, mensagem):
        """Adiciona mensagem ao log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.text_resultados.insert(tk.END, f"[{timestamp}] {mensagem}\n")
        self.text_resultados.see(tk.END)
        self.janela.update()

    def atualizar_tabela(self, dados=None):
        """Atualiza a tabela de resultados"""
        if dados is None:
            dados = self.dados_processados
            
        # Limpar tabela
        for item in self.tree_parcelamentos.get_children():
            self.tree_parcelamentos.delete(item)
        
        # Preencher tabela
        for row in dados:
            valores = (
                row['Nome_Empresa'],
                row['CNPJ'],
                row['Tipo'],
                row['Subtipo'],
                row['Conta'],
                row['Modalidade'],
                row['Status'],
                row['Detalhes'],
                f"R$ {row['Valor']:,.2f}" if row['Valor'] > 0 else "-",
                row['Arquivo']
            )
            self.tree_parcelamentos.insert("", tk.END, values=valores)
        
        self.label_contador.config(text=f"{len(dados)} resultados")

    def atualizar_dashboard(self):
        """Atualiza as estatísticas do dashboard"""
        if not self.dados_processados:
            return
            
        df = pd.DataFrame(self.dados_processados)
        
        # Estatísticas gerais
        total_empresas = df['Nome_Empresa'].nunique()
        total_parcelamentos = len(df)
        empresas_atraso = len(df[df['Detalhes'].str.contains('Parcelas em atraso', na=False)]['Nome_Empresa'].unique())
        valor_total = df['Valor'].sum()
        
        self.stats_labels['total_empresas'].config(text=str(total_empresas))
        self.stats_labels['total_parcelamentos'].config(text=str(total_parcelamentos))
        self.stats_labels['empresas_com_parcelas_atraso'].config(text=str(empresas_atraso))
        self.stats_labels['valor_total_suspenso'].config(text=f"R$ {valor_total:,.2f}")
        
        # Resumo por tipo
        for item in self.tree_resumo.get_children():
            self.tree_resumo.delete(item)
            
        resumo_tipo = df.groupby('Tipo').agg({
            'Nome_Empresa': 'nunique',
            'CNPJ': 'count'
        }).round(2)
        
        for tipo, row in resumo_tipo.iterrows():
            empresas = row['Nome_Empresa']
            quantidade = row['CNPJ']
            percentual = (quantidade / total_parcelamentos) * 100
            
            self.tree_resumo.insert("", tk.END, values=(
                tipo, quantidade, empresas, f"{percentual:.1f}%"
            ))

    def aplicar_filtros(self):
        """Aplica filtros na tabela"""
        if not self.dados_processados:
            return
            
        dados_filtrados = self.dados_processados.copy()
        
        # Filtro por empresa
        filtro_emp = self.filtro_empresa.get().lower()
        if filtro_emp:
            dados_filtrados = [d for d in dados_filtrados if filtro_emp in d['Nome_Empresa'].lower()]
        
        # Filtro por tipo
        filtro_tip = self.filtro_tipo.get()
        if filtro_tip != "Todos":
            dados_filtrados = [d for d in dados_filtrados if d['Tipo'] == filtro_tip]
        
        self.atualizar_tabela(dados_filtrados)

    def limpar_filtros(self):
        """Limpa todos os filtros"""
        self.filtro_empresa.delete(0, tk.END)
        self.filtro_tipo.set("Todos")
        self.atualizar_tabela()

    def limpar_resultados(self):
        """Limpa todos os resultados"""
        self.dados_processados = []
        self.text_resultados.delete(1.0, tk.END)
        self.atualizar_tabela()
        self.progress_var.set(0)
        self.status_label.config(text="Resultados limpos")

    def salvar_config(self):
        """Salva configuração atual"""
        config = {
            "pasta_pdfs": self.entrada_pasta_pdfs.get(),
            "excel_empresas": self.entrada_excel.get(),
            "pasta_saida": self.entrada_pasta_saida.get(),
            "incluir_detalhes_debitos": self.incluir_detalhes_debitos.get(),
            "agrupar_por_empresa": self.agrupar_por_empresa.get(),
            "salvar_backup_json": self.salvar_backup_json.get()
        }
        
        arquivo_config = filedialog.asksaveasfilename(
            title="Salvar configuração",
            defaultextension=".json",
            filetypes=[("JSON", "*.json")]
        )
        
        if arquivo_config:
            with open(arquivo_config, 'w') as f:
                json.dump(config, f, indent=2)
            messagebox.showinfo("Sucesso", "Configuração salva!")

    def exportar_filtrados(self):
        """Exporta dados atualmente filtrados"""
        items = self.tree_parcelamentos.get_children()
        if not items:
            messagebox.showwarning("Aviso", "Nenhum dado para exportar!")
            return
            
        arquivo = filedialog.asksaveasfilename(
            title="Exportar dados filtrados",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")]
        )
        
        if arquivo:
            # Coletar dados da tabela
            dados_exportar = []
            for item in items:
                valores = self.tree_parcelamentos.item(item)['values']
                dados_exportar.append(dict(zip(
                    ["Empresa", "CNPJ", "Tipo", "Subtipo", "Conta", "Modalidade", "Status", "Detalhes", "Valor", "Arquivo"],
                    valores
                )))
            
            df = pd.DataFrame(dados_exportar)
            
            if arquivo.endswith('.xlsx'):
                df.to_excel(arquivo, index=False)
            else:
                df.to_csv(arquivo, index=False, encoding='utf-8-sig')
                
            messagebox.showinfo("Sucesso", f"Dados exportados: {arquivo}")

    def copiar_selecionados(self):
        """Copia itens selecionados para clipboard"""
        selecionados = self.tree_parcelamentos.selection()
        if not selecionados:
            messagebox.showwarning("Aviso", "Selecione pelo menos um item!")
            return
        
        texto = ""
        for item in selecionados:
            valores = self.tree_parcelamentos.item(item)['values']
            texto += "\t".join(str(v) for v in valores) + "\n"
        
        self.janela.clipboard_clear()
        self.janela.clipboard_append(texto)
        messagebox.showinfo("Sucesso", f"{len(selecionados)} itens copiados!")

    def ordenar_coluna(self, coluna):
        """Ordena tabela por coluna"""
        # Implementar ordenação personalizada
        pass

    def mostrar_menu_contexto(self, event):
        """Mostra menu de contexto"""
        try:
            item = self.tree_parcelamentos.identify_row(event.y)
            if item:
                self.tree_parcelamentos.selection_set(item)
                self.menu_contexto.post(event.x_root, event.y_root)
        except:
            pass

    def copiar_cnpj(self):
        """Copia CNPJ selecionado"""
        item = self.tree_parcelamentos.selection()[0]
        cnpj = self.tree_parcelamentos.item(item)['values'][1]
        self.janela.clipboard_clear()
        self.janela.clipboard_append(cnpj)

    def copiar_linha(self):
        """Copia linha completa"""
        item = self.tree_parcelamentos.selection()[0]
        valores = self.tree_parcelamentos.item(item)['values']
        texto = "\t".join(str(v) for v in valores)
        self.janela.clipboard_clear()
        self.janela.clipboard_append(texto)

    def abrir_pdf(self):
        """Abre PDF correspondente"""
        item = self.tree_parcelamentos.selection()[0]
        arquivo = self.tree_parcelamentos.item(item)['values'][-1]
        pasta_pdfs = self.entrada_pasta_pdfs.get()
        caminho_pdf = os.path.join(pasta_pdfs, arquivo)
        
        if os.path.exists(caminho_pdf):
            os.startfile(caminho_pdf)  # Windows
        else:
            messagebox.showerror("Erro", "Arquivo PDF não encontrado!")

    def run(self):
        self.janela.mainloop()

# Executar aplicação
if __name__ == "__main__":
    app = AnalisadorParcelamentos()
    app.run()