import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from datetime import datetime
import numpy as np

class AnalisadorVendas:
    def __init__(self, root):
        self.root = root
        self.root.title("Analisador de Queda de Vendas")
        self.root.geometry("1200x800")
        
        self.df = None
        self.produtos_queda = None
        
        # Frame principal
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar grid
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # Frame de controles
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # Botão importar
        ttk.Button(control_frame, text="Importar CSV", command=self.importar_csv).pack(side=tk.LEFT, padx=5)
        
        # Botão analisar
        self.btn_analisar = ttk.Button(control_frame, text="Analisar Queda de Vendas", 
                                       command=self.analisar_vendas, state=tk.DISABLED)
        self.btn_analisar.pack(side=tk.LEFT, padx=5)
        
        # Botão exportar
        self.btn_exportar = ttk.Button(control_frame, text="Exportar Resultados", 
                                       command=self.exportar_resultados, state=tk.DISABLED)
        self.btn_exportar.pack(side=tk.LEFT, padx=5)
        
        # Label de status
        self.status_label = ttk.Label(control_frame, text="Nenhum arquivo importado")
        self.status_label.pack(side=tk.LEFT, padx=10)
        
        # Notebook (abas) para lista e gráfico
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Aba 1: Lista de produtos
        tab_lista = ttk.Frame(self.notebook)
        self.notebook.add(tab_lista, text="Ranking de Produtos")
        tab_lista.columnconfigure(0, weight=1)
        tab_lista.rowconfigure(0, weight=1)
        
        # Frame de filtros
        filter_frame = ttk.LabelFrame(tab_lista, text="Filtros", padding="5")
        filter_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=5, pady=5)
        filter_frame.columnconfigure(1, weight=1)
        filter_frame.columnconfigure(3, weight=1)
        filter_frame.columnconfigure(5, weight=1)
        
        # Filtros principais
        ttk.Label(filter_frame, text="Produto:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.filter_produto = ttk.Entry(filter_frame, width=20)
        self.filter_produto.grid(row=0, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        self.filter_produto.bind('<KeyRelease>', lambda e: self.aplicar_filtros())
        
        ttk.Label(filter_frame, text="Linha:").grid(row=0, column=2, padx=5, pady=2, sticky=tk.W)
        self.filter_linha = ttk.Entry(filter_frame, width=20)
        self.filter_linha.grid(row=0, column=3, padx=5, pady=2, sticky=(tk.W, tk.E))
        self.filter_linha.bind('<KeyRelease>', lambda e: self.aplicar_filtros())
        
        ttk.Label(filter_frame, text="Mês Início Queda:").grid(row=0, column=4, padx=5, pady=2, sticky=tk.W)
        self.filter_mes_queda = ttk.Entry(filter_frame, width=15)
        self.filter_mes_queda.grid(row=0, column=5, padx=5, pady=2, sticky=(tk.W, tk.E))
        self.filter_mes_queda.bind('<KeyRelease>', lambda e: self.aplicar_filtros())
        
        ttk.Label(filter_frame, text="Última Venda:").grid(row=1, column=0, padx=5, pady=2, sticky=tk.W)
        self.filter_ultima_venda = ttk.Entry(filter_frame, width=20)
        self.filter_ultima_venda.grid(row=1, column=1, padx=5, pady=2, sticky=(tk.W, tk.E))
        self.filter_ultima_venda.bind('<KeyRelease>', lambda e: self.aplicar_filtros())
        
        ttk.Label(filter_frame, text="Mês Mais Vendas:").grid(row=1, column=2, padx=5, pady=2, sticky=tk.W)
        self.filter_mes_mais_vendas = ttk.Entry(filter_frame, width=20)
        self.filter_mes_mais_vendas.grid(row=1, column=3, padx=5, pady=2, sticky=(tk.W, tk.E))
        self.filter_mes_mais_vendas.bind('<KeyRelease>', lambda e: self.aplicar_filtros())
        
        # Botão limpar filtros
        ttk.Button(filter_frame, text="Limpar Filtros", command=self.limpar_filtros).grid(row=1, column=4, columnspan=2, padx=5, pady=2)
        
        # Frame da lista
        list_frame = ttk.Frame(tab_lista, padding="5")
        list_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        tab_lista.rowconfigure(1, weight=1)
        
        # Treeview para lista
        columns = ('Rank', 'Produto', 'Linha', 'Quantidade Vendidas', 'Mês Início Queda', 
                  'Var. Preço %', 'Última Venda', 'Mês Mais Vendas', 'Qtd Mês Mais Vendas', 
                  'Mês Menos Vendas', 'Qtd Último Mês', 'Mês -2', 'Qtd Mês -2', 'Mês -1', 'Qtd Mês -1', 
                  'Mês Atual', 'Qtd Mês Atual', 'Média Vendas')
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=20)
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=90, anchor=tk.CENTER)
        
        self.tree.column('Produto', width=200, anchor=tk.W)
        self.tree.column('Linha', width=150, anchor=tk.W)
        self.tree.column('Quantidade Vendidas', width=130, anchor=tk.CENTER)
        self.tree.column('Mês Início Queda', width=110, anchor=tk.CENTER)
        self.tree.column('Var. Preço %', width=90, anchor=tk.CENTER)
        self.tree.column('Última Venda', width=100, anchor=tk.CENTER)
        self.tree.column('Mês Mais Vendas', width=110, anchor=tk.CENTER)
        self.tree.column('Qtd Mês Mais Vendas', width=120, anchor=tk.CENTER)
        self.tree.column('Mês Menos Vendas', width=110, anchor=tk.CENTER)
        self.tree.column('Qtd Último Mês', width=110, anchor=tk.CENTER)
        self.tree.column('Mês -2', width=90, anchor=tk.CENTER)
        self.tree.column('Qtd Mês -2', width=90, anchor=tk.CENTER)
        self.tree.column('Mês -1', width=90, anchor=tk.CENTER)
        self.tree.column('Qtd Mês -1', width=90, anchor=tk.CENTER)
        self.tree.column('Mês Atual', width=90, anchor=tk.CENTER)
        self.tree.column('Qtd Mês Atual', width=90, anchor=tk.CENTER)
        self.tree.column('Média Vendas', width=90, anchor=tk.CENTER)
        
        # Scrollbar vertical
        scrollbar_list_vertical = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar_list_vertical.set)
        
        # Scrollbar horizontal
        scrollbar_list_horizontal = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(xscrollcommand=scrollbar_list_horizontal.set)
        
        # Grid layout
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar_list_vertical.grid(row=0, column=1, sticky=(tk.N, tk.S))
        scrollbar_list_horizontal.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Armazenar dados originais para filtro
        self.produtos_queda_original = None
        
        # Aba 2: Gráfico
        tab_grafico = ttk.Frame(self.notebook)
        self.notebook.add(tab_grafico, text="Gráfico de Queda")
        tab_grafico.columnconfigure(0, weight=1)
        tab_grafico.rowconfigure(0, weight=1)
        
        # Frame do gráfico
        self.chart_frame = ttk.Frame(tab_grafico, padding="5")
        self.chart_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.chart_frame.columnconfigure(0, weight=1)
        self.chart_frame.rowconfigure(0, weight=1)
        
        self.fig = None
        self.canvas = None
        
        # Aba 3: Gráfico de Pizza por Linha
        tab_pizza = ttk.Frame(self.notebook)
        self.notebook.add(tab_pizza, text="Queda por Linha")
        tab_pizza.columnconfigure(0, weight=1)
        tab_pizza.rowconfigure(0, weight=1)
        
        # Frame do gráfico de pizza
        self.chart_pizza_frame = ttk.Frame(tab_pizza, padding="5")
        self.chart_pizza_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.chart_pizza_frame.columnconfigure(0, weight=1)
        self.chart_pizza_frame.rowconfigure(0, weight=1)
        
        self.fig_pizza = None
        self.canvas_pizza = None
        
    def importar_csv(self):
        """Importa arquivo CSV"""
        arquivo = filedialog.askopenfilename(
            title="Selecione o arquivo CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if arquivo:
            try:
                # Tentar diferentes encodings - formato do arquivo vendascachoeira.csv
                encodings = ['latin-1', 'cp1252', 'iso-8859-1', 'utf-8']
                self.df = None
                
                # Primeiro, encontrar a linha do cabeçalho lendo o arquivo
                linha_cabecalho = None
                encoding_usado = None
                
                for encoding in encodings:
                    try:
                        with open(arquivo, 'r', encoding=encoding) as f:
                            linhas = f.readlines()
                            for i, linha in enumerate(linhas[:20]):  # Verificar apenas as primeiras 20 linhas
                                linha_limpa = linha.strip()
                                # Ignorar linhas vazias ou que são apenas aspas
                                if not linha_limpa or linha_limpa == '"':
                                    continue
                                # Procurar pela linha que contém "CodFilial" e "DataVenda"
                                if 'CodFilial' in linha and 'DataVenda' in linha:
                                    linha_cabecalho = i
                                    encoding_usado = encoding
                                    break
                        if linha_cabecalho is not None:
                            break
                    except:
                        continue
                
                if linha_cabecalho is None:
                    # Se não encontrou, usar linha 7 (índice 7 = linha 8 do arquivo)
                    # Mas verificar se a linha 7 é apenas aspas, então usar linha 8
                    linha_cabecalho = 7
                    encoding_usado = encodings[0]
                
                # Ler a linha do cabeçalho para obter os nomes das colunas
                nomes_colunas = None
                if encoding_usado and linha_cabecalho is not None:
                    try:
                        with open(arquivo, 'r', encoding=encoding_usado) as f:
                            linhas = f.readlines()
                            if linha_cabecalho < len(linhas):
                                linha_header = linhas[linha_cabecalho].strip()
                                # Remover aspas se houver
                                linha_header = linha_header.strip('"')
                                # Dividir por vírgula
                                nomes_colunas = [col.strip().strip('"') for col in linha_header.split(',')]
                    except:
                        pass
                
                # Agora ler o CSV usando header para especificar a linha do cabeçalho
                self.df = None
                for encoding in [encoding_usado] if encoding_usado else encodings:
                    try:
                        # Se temos os nomes das colunas, usar skiprows e names
                        if nomes_colunas and len(nomes_colunas) > 5:
                            temp_df = pd.read_csv(
                                arquivo, 
                                encoding=encoding, 
                                sep=',', 
                                decimal=',',
                                skiprows=linha_cabecalho + 1,  # Pular a linha do cabeçalho também
                                names=nomes_colunas,
                                quotechar='"',
                                on_bad_lines='skip',
                                engine='python'
                            )
                        else:
                            # Tentar com header
                            temp_df = pd.read_csv(
                                arquivo, 
                                encoding=encoding, 
                                sep=',', 
                                decimal=',',
                                header=linha_cabecalho,  # Usar header para especificar a linha do cabeçalho
                                quotechar='"',
                                on_bad_lines='skip',
                                engine='python'
                            )
                        
                        # Verificar se as colunas esperadas estão presentes
                        colunas_str = ' '.join(temp_df.columns.astype(str))
                        if 'CodFilial' in colunas_str or 'DataVenda' in colunas_str:
                            self.df = temp_df
                            break
                    except Exception as e:
                        continue
                
                # Se ainda não encontrou, tentar com skiprows simples
                if self.df is None or len(self.df) == 0:
                    for encoding in encodings:
                        try:
                            temp_df = pd.read_csv(
                                arquivo, 
                                encoding=encoding, 
                                sep=',', 
                                decimal=',',
                                skiprows=linha_cabecalho + 1 if linha_cabecalho else 8,
                                quotechar='"',
                                on_bad_lines='skip',
                                engine='python'
                            )
                            colunas_str = ' '.join(temp_df.columns.astype(str))
                            if 'CodFilial' in colunas_str or 'DataVenda' in colunas_str:
                                self.df = temp_df
                                break
                        except:
                            continue
                
                if self.df is None or len(self.df) == 0:
                    messagebox.showerror("Erro", "Não foi possível ler o arquivo CSV. Verifique o formato do arquivo.")
                    return
                
                # Normalizar nomes das colunas (remover espaços e caracteres especiais)
                self.df.columns = self.df.columns.str.strip()
                
                # Mapear nomes de colunas possíveis
                coluna_data = None
                coluna_produto = None
                coluna_quantidade = None
                
                # Procurar coluna de data (prioridade para nomes exatos)
                for col in self.df.columns:
                    col_clean = col.strip()
                    if col_clean == 'DataVenda':
                        coluna_data = col
                        break
                    elif 'DataVenda' in col_clean:
                        coluna_data = col
                        break
                
                # Se não encontrou, procurar por "Data"
                if not coluna_data:
                    for col in self.df.columns:
                        if 'Data' in col and 'Venda' in col:
                            coluna_data = col
                            break
                
                # Procurar coluna de produto (prioridade para nomes exatos)
                # Normalizar para buscar sem considerar acentos
                for col in self.df.columns:
                    col_clean = col.strip()
                    # Buscar por diferentes variações
                    if 'DescriçãoProduto' in col_clean or 'DescricaoProduto' in col_clean:
                        coluna_produto = col
                        break
                    elif ('Descrição' in col_clean or 'Descricao' in col_clean) and 'Produto' in col_clean:
                        coluna_produto = col
                        break
                    elif 'Produto' in col_clean and len(col_clean) > 5:  # Evitar falsos positivos
                        coluna_produto = col
                        break
                
                # Procurar coluna de quantidade (prioridade para QtdeItem)
                for col in self.df.columns:
                    col_clean = col.strip()
                    if col_clean == 'QtdeItem':
                        coluna_quantidade = col
                        break
                    elif 'QtdeItem' in col_clean:
                        coluna_quantidade = col
                        break
                    elif 'QtdItem' in col_clean:
                        coluna_quantidade = col
                        break
                
                if not coluna_data or not coluna_produto or not coluna_quantidade:
                    colunas_faltando = []
                    if not coluna_data:
                        colunas_faltando.append("DataVenda")
                    if not coluna_produto:
                        colunas_faltando.append("DescriçãoProduto")
                    if not coluna_quantidade:
                        colunas_faltando.append("QtdeItem")
                    
                    messagebox.showerror("Erro", 
                        f"Colunas necessárias não encontradas: {', '.join(colunas_faltando)}\n\n"
                        f"Colunas disponíveis no arquivo:\n{', '.join(self.df.columns)}")
                    return
                
                # Renomear colunas para facilitar
                self.df = self.df.rename(columns={
                    coluna_data: 'DataVenda',
                    coluna_produto: 'DescricaoProduto',
                    coluna_quantidade: 'QtdItem'
                })
                
                # Preencher campos vazios com valores anteriores (formato agrupado do arquivo)
                # Preencher CodFilial, NomeFilial, DescricaoProduto, EAN
                colunas_preencher = ['CodFilial', 'NomeFilial', 'DescricaoProduto', 'EAN', 'Linha']
                for col in colunas_preencher:
                    if col in self.df.columns:
                        # Converter para string e tratar valores vazios
                        self.df[col] = self.df[col].astype(str).replace(['nan', 'None', ''], np.nan)
                        self.df[col] = self.df[col].ffill()
                
                # Remover linhas onde tanto DataVenda quanto DescricaoProduto estão vazios
                self.df = self.df.dropna(subset=['DataVenda', 'DescricaoProduto'], how='all')
                
                # Remover linhas onde DescricaoProduto está vazio ou é 'nan'
                self.df = self.df[self.df['DescricaoProduto'].notna()]
                self.df = self.df[self.df['DescricaoProduto'].astype(str).str.strip() != '']
                self.df = self.df[self.df['DescricaoProduto'].astype(str).str.lower() != 'nan']
                
                # Converter DataVenda para datetime (pode estar com ou sem aspas)
                self.df['DataVenda'] = self.df['DataVenda'].astype(str).str.strip().str.replace('"', '')
                self.df['DataVenda'] = pd.to_datetime(self.df['DataVenda'], format='%d/%m/%Y', errors='coerce')
                
                # Remover linhas com data inválida
                self.df = self.df.dropna(subset=['DataVenda'])
                
                # Limpar e converter quantidade (pode estar com vírgula como decimal)
                if self.df['QtdItem'].dtype == 'object':
                    self.df['QtdItem'] = self.df['QtdItem'].astype(str).str.replace(',', '.').str.replace('"', '')
                self.df['QtdItem'] = pd.to_numeric(self.df['QtdItem'], errors='coerce').fillna(0)
                
                # Remover linhas com quantidade zero ou inválida (opcional - pode comentar se quiser manter)
                # self.df = self.df[self.df['QtdItem'] > 0]
                
                # Limpar nome do produto
                if self.df['DescricaoProduto'].dtype == 'object':
                    self.df['DescricaoProduto'] = self.df['DescricaoProduto'].astype(str).str.strip()
                
                self.status_label.config(text=f"Arquivo importado: {len(self.df)} registros")
                self.btn_analisar.config(state=tk.NORMAL)
                
                messagebox.showinfo("Sucesso", f"Arquivo importado com sucesso!\n{len(self.df)} registros carregados.")
                
            except Exception as e:
                import traceback
                messagebox.showerror("Erro", f"Erro ao importar arquivo:\n{str(e)}\n\n{traceback.format_exc()}")
    
    def identificar_mes_inicio_queda(self, df_sorted, produto):
        """Identifica o mês em que começou a queda de vendas do produto"""
        try:
            # Filtrar dados do produto
            df_produto = df_sorted[df_sorted['DescricaoProduto'] == produto].copy()
            
            if len(df_produto) < 2:
                return 'N/A'
            
            # Agrupar por mês
            df_produto['AnoMes'] = df_produto['DataVenda'].dt.to_period('M')
            vendas_mes = df_produto.groupby('AnoMes')['QtdItem'].sum().reset_index()
            vendas_mes = vendas_mes.sort_values('AnoMes')
            
            if len(vendas_mes) < 2:
                return 'N/A'
            
            # Calcular variação mês a mês
            vendas_mes['Variacao'] = vendas_mes['QtdItem'].pct_change() * 100
            
            # Encontrar primeiro mês com queda significativa (mais de 10% de queda)
            for i in range(1, len(vendas_mes)):
                if vendas_mes.iloc[i]['Variacao'] < -10:  # Queda de mais de 10%
                    mes_queda = vendas_mes.iloc[i]['AnoMes']
                    return f"{mes_queda.strftime('%m/%Y')}"
            
            # Se não encontrou queda significativa, retornar último mês com queda
            for i in range(1, len(vendas_mes)):
                if vendas_mes.iloc[i]['Variacao'] < 0:
                    mes_queda = vendas_mes.iloc[i]['AnoMes']
                    return f"{mes_queda.strftime('%m/%Y')}"
            
            return 'N/A'
        except:
            return 'N/A'
    
    def obter_ultima_venda(self, df_sorted, produto):
        """Obtém a data da última venda do produto"""
        try:
            df_produto = df_sorted[df_sorted['DescricaoProduto'] == produto].copy()
            if len(df_produto) == 0:
                return 'N/A'
            ultima_data = df_produto['DataVenda'].max()
            return ultima_data.strftime('%d/%m/%Y')
        except:
            return 'N/A'
    
    def obter_mes_mais_vendas(self, df_sorted, produto):
        """Obtém o mês com mais vendas do produto e a quantidade"""
        try:
            df_produto = df_sorted[df_sorted['DescricaoProduto'] == produto].copy()
            if len(df_produto) == 0:
                return ('N/A', 0)
            
            # Agrupar por mês
            df_produto['AnoMes'] = df_produto['DataVenda'].dt.to_period('M')
            vendas_mes = df_produto.groupby('AnoMes')['QtdItem'].sum().reset_index()
            vendas_mes = vendas_mes.sort_values('QtdItem', ascending=False)
            
            if len(vendas_mes) == 0:
                return ('N/A', 0)
            
            mes_mais_vendas = vendas_mes.iloc[0]['AnoMes']
            qtd_mais_vendas = vendas_mes.iloc[0]['QtdItem']
            return (f"{mes_mais_vendas.strftime('%m/%Y')}", qtd_mais_vendas)
        except:
            return ('N/A', 0)
    
    def obter_mes_menos_vendas(self, df_sorted, produto):
        """Obtém o mês com menos vendas do produto e a quantidade"""
        try:
            df_produto = df_sorted[df_sorted['DescricaoProduto'] == produto].copy()
            if len(df_produto) == 0:
                return ('N/A', 0)
            
            # Agrupar por mês
            df_produto['AnoMes'] = df_produto['DataVenda'].dt.to_period('M')
            vendas_mes = df_produto.groupby('AnoMes')['QtdItem'].sum().reset_index()
            vendas_mes = vendas_mes.sort_values('QtdItem', ascending=True)
            
            if len(vendas_mes) == 0:
                return ('N/A', 0)
            
            mes_menos_vendas = vendas_mes.iloc[0]['AnoMes']
            qtd_menos_vendas = vendas_mes.iloc[0]['QtdItem']
            return (f"{mes_menos_vendas.strftime('%m/%Y')}", qtd_menos_vendas)
        except:
            return ('N/A', 0)
    
    def obter_qtd_ultimo_mes(self, df_sorted, produto):
        """Obtém a quantidade de vendas do último mês do produto"""
        try:
            df_produto = df_sorted[df_sorted['DescricaoProduto'] == produto].copy()
            if len(df_produto) == 0:
                return 0
            
            # Agrupar por mês
            df_produto['AnoMes'] = df_produto['DataVenda'].dt.to_period('M')
            vendas_mes = df_produto.groupby('AnoMes')['QtdItem'].sum().reset_index()
            vendas_mes = vendas_mes.sort_values('AnoMes', ascending=False)
            
            if len(vendas_mes) == 0:
                return 0
            
            # Pegar o último mês
            qtd_ultimo_mes = vendas_mes.iloc[0]['QtdItem']
            return qtd_ultimo_mes
        except:
            return 0
    
    def obter_ultimos_3_meses(self, df_sorted, produto):
        """Obtém os últimos 3 meses e suas quantidades de vendas"""
        try:
            df_produto = df_sorted[df_sorted['DescricaoProduto'] == produto].copy()
            if len(df_produto) == 0:
                return (('N/A', 0), ('N/A', 0), ('N/A', 0))
            
            # Agrupar por mês
            df_produto['AnoMes'] = df_produto['DataVenda'].dt.to_period('M')
            vendas_mes = df_produto.groupby('AnoMes')['QtdItem'].sum().reset_index()
            vendas_mes = vendas_mes.sort_values('AnoMes', ascending=False)
            
            if len(vendas_mes) == 0:
                return (('N/A', 0), ('N/A', 0), ('N/A', 0))
            
            # Pegar os últimos 3 meses
            mes_atual = ('N/A', 0)
            mes_menos_1 = ('N/A', 0)
            mes_menos_2 = ('N/A', 0)
            
            if len(vendas_mes) >= 1:
                mes_atual = (f"{vendas_mes.iloc[0]['AnoMes'].strftime('%m/%Y')}", vendas_mes.iloc[0]['QtdItem'])
            if len(vendas_mes) >= 2:
                mes_menos_1 = (f"{vendas_mes.iloc[1]['AnoMes'].strftime('%m/%Y')}", vendas_mes.iloc[1]['QtdItem'])
            if len(vendas_mes) >= 3:
                mes_menos_2 = (f"{vendas_mes.iloc[2]['AnoMes'].strftime('%m/%Y')}", vendas_mes.iloc[2]['QtdItem'])
            
            return (mes_menos_2, mes_menos_1, mes_atual)
        except:
            return (('N/A', 0), ('N/A', 0), ('N/A', 0))
    
    def calcular_media_vendas(self, df_sorted, produto):
        """Calcula a média de vendas mensais do produto"""
        try:
            df_produto = df_sorted[df_sorted['DescricaoProduto'] == produto].copy()
            if len(df_produto) == 0:
                return 0
            
            # Agrupar por mês
            df_produto['AnoMes'] = df_produto['DataVenda'].dt.to_period('M')
            vendas_mes = df_produto.groupby('AnoMes')['QtdItem'].sum().reset_index()
            
            if len(vendas_mes) == 0:
                return 0
            
            media = vendas_mes['QtdItem'].mean()
            return media
        except:
            return 0
    
    def calcular_variacao_preco(self, df_periodo1, df_periodo2, produto):
        """Calcula a variação de preço do produto entre os dois períodos"""
        try:
            # Verificar se há coluna de preço (procurar por diferentes nomes possíveis)
            coluna_preco = None
            colunas_possiveis = ['VlrVenda', 'VlrLiqVenda', 'VlrVenda', 'VlrLiqVenda']
            
            # Verificar no dataframe original também
            if hasattr(self, 'df') and self.df is not None:
                for col in self.df.columns:
                    if 'Vlr' in col and ('Venda' in col or 'Liq' in col):
                        colunas_possiveis.append(col)
            
            for col in colunas_possiveis:
                if col in df_periodo1.columns:
                    coluna_preco = col
                    break
            
            if coluna_preco is None:
                return np.nan
            
            # Filtrar produto em cada período
            df_p1 = df_periodo1[df_periodo1['DescricaoProduto'] == produto].copy()
            df_p2 = df_periodo2[df_periodo2['DescricaoProduto'] == produto].copy()
            
            if len(df_p1) == 0 or len(df_p2) == 0:
                return np.nan
            
            # Converter preço para numérico (pode ter vírgula como decimal)
            if df_p1[coluna_preco].dtype == 'object':
                df_p1[coluna_preco] = df_p1[coluna_preco].astype(str).str.replace(',', '.').str.replace('"', '').str.strip()
            if df_p2[coluna_preco].dtype == 'object':
                df_p2[coluna_preco] = df_p2[coluna_preco].astype(str).str.replace(',', '.').str.replace('"', '').str.strip()
            
            df_p1[coluna_preco] = pd.to_numeric(df_p1[coluna_preco], errors='coerce')
            df_p2[coluna_preco] = pd.to_numeric(df_p2[coluna_preco], errors='coerce')
            
            # Remover valores NaN
            df_p1 = df_p1.dropna(subset=[coluna_preco])
            df_p2 = df_p2.dropna(subset=[coluna_preco])
            
            if len(df_p1) == 0 or len(df_p2) == 0:
                return np.nan
            
            # Calcular preço médio ponderado por quantidade em cada período
            df_p1['PrecoQtd'] = df_p1[coluna_preco] * df_p1['QtdItem']
            df_p2['PrecoQtd'] = df_p2[coluna_preco] * df_p2['QtdItem']
            
            qtd_total_p1 = df_p1['QtdItem'].sum()
            qtd_total_p2 = df_p2['QtdItem'].sum()
            
            if qtd_total_p1 == 0 or qtd_total_p2 == 0:
                return np.nan
            
            preco_medio_p1 = df_p1['PrecoQtd'].sum() / qtd_total_p1
            preco_medio_p2 = df_p2['PrecoQtd'].sum() / qtd_total_p2
            
            if preco_medio_p1 == 0:
                return np.nan
            
            # Calcular variação percentual
            variacao = ((preco_medio_p2 - preco_medio_p1) / preco_medio_p1) * 100
            return variacao  # Retornar número para formatação na lista
        except Exception as e:
            return np.nan
    
    def analisar_vendas(self):
        """Analisa queda de vendas comparando períodos"""
        if self.df is None or len(self.df) == 0:
            messagebox.showwarning("Aviso", "Nenhum dado disponível para análise")
            return
        
        try:
            # Ordenar por data
            df_sorted = self.df.sort_values('DataVenda')
            
            # Determinar período total
            data_min = df_sorted['DataVenda'].min()
            data_max = df_sorted['DataVenda'].max()
            
            # Dividir em dois períodos (primeiros 3 meses vs últimos 3 meses)
            periodo_total_dias = (data_max - data_min).days
            data_meio = data_min + pd.Timedelta(days=periodo_total_dias / 2)
            
            # Primeiro período (inicial)
            df_periodo1 = df_sorted[df_sorted['DataVenda'] < data_meio].copy()
            # Segundo período (final)
            df_periodo2 = df_sorted[df_sorted['DataVenda'] >= data_meio].copy()
            
            # Agrupar por produto em cada período (incluindo linha)
            periodo1 = df_periodo1.groupby('DescricaoProduto').agg({
                'QtdItem': 'sum',
                'Linha': 'first'  # Pegar a primeira linha do produto
            }).reset_index()
            periodo1.columns = ['Produto', 'Qtd_Periodo1', 'Linha']
            
            periodo2 = df_periodo2.groupby('DescricaoProduto')['QtdItem'].sum().reset_index()
            periodo2.columns = ['Produto', 'Qtd_Periodo2']
            
            # Merge dos dois períodos
            comparacao = periodo1.merge(periodo2, on='Produto', how='outer').fillna(0)
            
            # Garantir que Linha está presente
            if 'Linha' not in comparacao.columns:
                # Se não tiver coluna Linha, buscar do dataframe original
                linhas_produtos = df_sorted.groupby('DescricaoProduto')['Linha'].first().reset_index()
                linhas_produtos.columns = ['Produto', 'Linha']
                comparacao = comparacao.merge(linhas_produtos, on='Produto', how='left')
                comparacao['Linha'] = comparacao['Linha'].fillna('N/A')
            
            # Calcular quantidade total vendida diretamente do dataframe completo para garantir precisão
            qtd_total_completa = df_sorted.groupby('DescricaoProduto')['QtdItem'].sum().reset_index()
            qtd_total_completa.columns = ['Produto', 'Quantidade_Vendidas']
            comparacao = comparacao.merge(qtd_total_completa, on='Produto', how='left')
            comparacao['Quantidade_Vendidas'] = comparacao['Quantidade_Vendidas'].fillna(0)
            
            # Calcular variação percentual (melhorada)
            # Considerar queda apenas se:
            # 1. Teve vendas no período 1 E teve menos vendas no período 2, OU
            # 2. Teve vendas consistentes e depois parou completamente
            comparacao['Variação_%'] = ((comparacao['Qtd_Periodo2'] - comparacao['Qtd_Periodo1']) / 
                                       comparacao['Qtd_Periodo1'].replace(0, np.nan) * 100).fillna(0)
            
            # Analisar tendência mês a mês para detectar quedas reais
            def verificar_queda_real(produto):
                df_prod = df_sorted[df_sorted['DescricaoProduto'] == produto].copy()
                if len(df_prod) < 2:
                    return False
                
                # Agrupar por mês
                df_prod['AnoMes'] = df_prod['DataVenda'].dt.to_period('M')
                vendas_mes = df_prod.groupby('AnoMes')['QtdItem'].sum().reset_index()
                vendas_mes = vendas_mes.sort_values('AnoMes')
                
                if len(vendas_mes) < 2:
                    return False
                
                # Verificar se houve tendência de queda:
                # 1. Se o último mês tem menos vendas que o primeiro mês com vendas significativas
                # 2. Ou se parou completamente de vender
                
                # Encontrar primeiro mês com vendas significativas (>0)
                primeiro_mes_com_vendas = vendas_mes[vendas_mes['QtdItem'] > 0]
                if len(primeiro_mes_com_vendas) == 0:
                    return False
                
                qtd_primeiro_mes = primeiro_mes_com_vendas.iloc[0]['QtdItem']
                qtd_ultimo_mes = vendas_mes.iloc[-1]['QtdItem']
                
                # Verificar se parou de vender
                if qtd_ultimo_mes == 0 and qtd_primeiro_mes > 0:
                    return True
                
                # Verificar se houve queda significativa (último mês < 50% do primeiro mês com vendas)
                if qtd_ultimo_mes > 0 and qtd_primeiro_mes > 0:
                    if qtd_ultimo_mes < (qtd_primeiro_mes * 0.5):
                        return True
                
                # Verificar tendência: se os últimos 2 meses têm menos vendas que os primeiros 2 meses
                if len(vendas_mes) >= 4:
                    primeiros_2_meses = vendas_mes.head(2)['QtdItem'].sum()
                    ultimos_2_meses = vendas_mes.tail(2)['QtdItem'].sum()
                    if primeiros_2_meses > 0 and ultimos_2_meses < (primeiros_2_meses * 0.5):
                        return True
                
                return False
            
            # Aplicar verificação de queda real
            comparacao['Tem_Queda_Real'] = comparacao['Produto'].apply(verificar_queda_real)
            
            # Filtrar produtos com queda real OU que pararam de vender no período final
            produtos_queda = comparacao[
                (comparacao['Tem_Queda_Real']) | 
                ((comparacao['Qtd_Periodo1'] > 0) & (comparacao['Qtd_Periodo2'] == 0)) |
                (comparacao['Variação_%'] < -20)  # Queda de mais de 20%
            ].copy()
            
            # Recalcular variação para produtos que realmente tiveram queda
            produtos_queda['Variação_%'] = ((produtos_queda['Qtd_Periodo2'] - produtos_queda['Qtd_Periodo1']) / 
                                           produtos_queda['Qtd_Periodo1'].replace(0, np.nan) * 100).fillna(-100)
            
            # Analisar mês a mês para identificar quando começou a queda
            produtos_queda['Mes_Inicio_Queda'] = produtos_queda['Produto'].apply(
                lambda p: self.identificar_mes_inicio_queda(df_sorted, p)
            )
            
            # Calcular variação de preço se houver coluna de preço
            produtos_queda['Variacao_Preco_%'] = produtos_queda['Produto'].apply(
                lambda p: self.calcular_variacao_preco(df_periodo1, df_periodo2, p)
            )
            
            # Calcular data da última venda
            produtos_queda['Ultima_Venda'] = produtos_queda['Produto'].apply(
                lambda p: self.obter_ultima_venda(df_sorted, p)
            )
            
            # Calcular mês com mais vendas e quantidade
            resultado_mes_mais_vendas = produtos_queda['Produto'].apply(
                lambda p: self.obter_mes_mais_vendas(df_sorted, p)
            )
            produtos_queda['Mes_Mais_Vendas'] = resultado_mes_mais_vendas.apply(lambda x: x[0])
            produtos_queda['Qtd_Mes_Mais_Vendas'] = resultado_mes_mais_vendas.apply(lambda x: x[1])
            
            # Calcular mês com menos vendas e quantidade
            resultado_mes_menos_vendas = produtos_queda['Produto'].apply(
                lambda p: self.obter_mes_menos_vendas(df_sorted, p)
            )
            produtos_queda['Mes_Menos_Vendas'] = resultado_mes_menos_vendas.apply(lambda x: x[0])
            produtos_queda['Qtd_Mes_Menos_Vendas'] = resultado_mes_menos_vendas.apply(lambda x: x[1])
            
            # Calcular quantidade do último mês
            produtos_queda['Qtd_Ultimo_Mes'] = produtos_queda['Produto'].apply(
                lambda p: self.obter_qtd_ultimo_mes(df_sorted, p)
            )
            
            # Calcular últimos 3 meses
            resultado_ultimos_3_meses = produtos_queda['Produto'].apply(
                lambda p: self.obter_ultimos_3_meses(df_sorted, p)
            )
            produtos_queda['Mes_Menos_2'] = resultado_ultimos_3_meses.apply(lambda x: x[0][0])
            produtos_queda['Qtd_Mes_Menos_2'] = resultado_ultimos_3_meses.apply(lambda x: x[0][1])
            produtos_queda['Mes_Menos_1'] = resultado_ultimos_3_meses.apply(lambda x: x[1][0])
            produtos_queda['Qtd_Mes_Menos_1'] = resultado_ultimos_3_meses.apply(lambda x: x[1][1])
            produtos_queda['Mes_Atual'] = resultado_ultimos_3_meses.apply(lambda x: x[2][0])
            produtos_queda['Qtd_Mes_Atual'] = resultado_ultimos_3_meses.apply(lambda x: x[2][1])
            
            # Calcular média de vendas
            produtos_queda['Media_Vendas'] = produtos_queda['Produto'].apply(
                lambda p: self.calcular_media_vendas(df_sorted, p)
            )
            
            # Identificar produtos que pararam de vender (sem vendas no período final)
            produtos_queda['Parou_Vender'] = produtos_queda['Qtd_Periodo2'] == 0
            
            # Calcular score de ranking: priorizar produtos que pararam de vender e tiveram mais vendas
            # Score = (Parou de vender? 1000000 : 0) + Quantidade_Vendidas
            # Isso garante que produtos que pararam apareçam primeiro, ordenados por quantidade vendida
            produtos_queda['Score_Ranking'] = (
                produtos_queda['Parou_Vender'].astype(int) * 1000000 + 
                produtos_queda['Quantidade_Vendidas']
            )
            
            # Ordenar: primeiro os que pararam de vender (ordenados por maior quantidade vendida)
            # depois os que ainda vendem mas tiveram queda (ordenados por maior quantidade vendida)
            produtos_queda = produtos_queda.sort_values('Score_Ranking', ascending=False)
            
            # Pegar os 100 primeiros de cada Linha
            produtos_queda = produtos_queda.groupby('Linha').head(100).reset_index(drop=True)
            
            # Adicionar ranking
            produtos_queda['Rank'] = range(1, len(produtos_queda) + 1)
            
            # Garantir que Linha está presente e preencher se necessário
            if 'Linha' not in produtos_queda.columns:
                linhas_produtos = df_sorted.groupby('DescricaoProduto')['Linha'].first().reset_index()
                linhas_produtos.columns = ['Produto', 'Linha']
                produtos_queda = produtos_queda.merge(linhas_produtos, on='Produto', how='left')
                produtos_queda['Linha'] = produtos_queda['Linha'].fillna('N/A')
            
            self.produtos_queda = produtos_queda
            self.produtos_queda_original = produtos_queda.copy()  # Guardar cópia para filtros
            
            # Atualizar lista
            self.atualizar_lista()
            
            # Atualizar gráfico
            self.atualizar_grafico()
            
            # Atualizar gráfico de pizza
            self.atualizar_grafico_pizza()
            
            # Habilitar botão de exportar
            self.btn_exportar.config(state=tk.NORMAL)
            
            self.status_label.config(text=f"Análise concluída: {len(produtos_queda)} produtos com queda")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao analisar vendas:\n{str(e)}")
    
    def atualizar_lista(self):
        """Atualiza a lista de produtos"""
        # Limpar lista atual
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        if self.produtos_queda is None or len(self.produtos_queda) == 0:
            return
        
        # Adicionar itens
        for _, row in self.produtos_queda.iterrows():
            mes_queda = row.get('Mes_Inicio_Queda', 'N/A')
            var_preco = row.get('Variacao_Preco_%', np.nan)
            ultima_venda = row.get('Ultima_Venda', 'N/A')
            mes_mais_vendas = row.get('Mes_Mais_Vendas', 'N/A')
            qtd_mes_mais_vendas = row.get('Qtd_Mes_Mais_Vendas', 0)
            mes_menos_vendas = row.get('Mes_Menos_Vendas', 'N/A')
            qtd_ultimo_mes = row.get('Qtd_Ultimo_Mes', 0)
            mes_menos_2 = row.get('Mes_Menos_2', 'N/A')
            qtd_mes_menos_2 = row.get('Qtd_Mes_Menos_2', 0)
            mes_menos_1 = row.get('Mes_Menos_1', 'N/A')
            qtd_mes_menos_1 = row.get('Qtd_Mes_Menos_1', 0)
            mes_atual = row.get('Mes_Atual', 'N/A')
            qtd_mes_atual = row.get('Qtd_Mes_Atual', 0)
            media_vendas = row.get('Media_Vendas', 0)
            linha = row.get('Linha', 'N/A')
            
            # Formatar variação de preço
            if pd.isna(var_preco) or var_preco == 'N/A':
                var_preco_str = 'N/A'
            else:
                var_preco_str = f"{var_preco:.1f}%"
            
            quantidade_vendidas = row.get('Quantidade_Vendidas', row.get('Qtd_Total', 0))
            
            self.tree.insert('', 'end', values=(
                int(row['Rank']),
                row['Produto'],
                str(linha),
                f"{quantidade_vendidas:.0f}",
                str(mes_queda),
                var_preco_str,
                str(ultima_venda),
                str(mes_mais_vendas),
                f"{qtd_mes_mais_vendas:.0f}",
                str(mes_menos_vendas),
                f"{qtd_ultimo_mes:.0f}",
                str(mes_menos_2),
                f"{qtd_mes_menos_2:.0f}",
                str(mes_menos_1),
                f"{qtd_mes_menos_1:.0f}",
                str(mes_atual),
                f"{qtd_mes_atual:.0f}",
                f"{media_vendas:.1f}"
            ))
    
    def aplicar_filtros(self):
        """Aplica os filtros na lista de produtos"""
        if self.produtos_queda_original is None or len(self.produtos_queda_original) == 0:
            return
        
        # Obter valores dos filtros
        filtro_produto = self.filter_produto.get().strip().lower()
        filtro_linha = self.filter_linha.get().strip().lower()
        filtro_mes_queda = self.filter_mes_queda.get().strip().lower()
        filtro_ultima_venda = self.filter_ultima_venda.get().strip().lower()
        filtro_mes_mais_vendas = self.filter_mes_mais_vendas.get().strip().lower()
        
        # Aplicar filtros
        df_filtrado = self.produtos_queda_original.copy()
        
        if filtro_produto:
            df_filtrado = df_filtrado[
                df_filtrado['Produto'].astype(str).str.lower().str.contains(filtro_produto, na=False)
            ]
        
        if filtro_linha:
            df_filtrado = df_filtrado[
                df_filtrado['Linha'].astype(str).str.lower().str.contains(filtro_linha, na=False)
            ]
        
        if filtro_mes_queda:
            df_filtrado = df_filtrado[
                df_filtrado['Mes_Inicio_Queda'].astype(str).str.lower().str.contains(filtro_mes_queda, na=False)
            ]
        
        if filtro_ultima_venda:
            df_filtrado = df_filtrado[
                df_filtrado['Ultima_Venda'].astype(str).str.lower().str.contains(filtro_ultima_venda, na=False)
            ]
        
        if filtro_mes_mais_vendas:
            df_filtrado = df_filtrado[
                df_filtrado['Mes_Mais_Vendas'].astype(str).str.lower().str.contains(filtro_mes_mais_vendas, na=False)
            ]
        
        # Atualizar ranking
        df_filtrado = df_filtrado.sort_values('Variação_%')
        df_filtrado['Rank'] = range(1, len(df_filtrado) + 1)
        
        self.produtos_queda = df_filtrado
        self.atualizar_lista()
    
    def limpar_filtros(self):
        """Limpa todos os filtros"""
        self.filter_produto.delete(0, tk.END)
        self.filter_linha.delete(0, tk.END)
        self.filter_mes_queda.delete(0, tk.END)
        self.filter_ultima_venda.delete(0, tk.END)
        self.filter_mes_mais_vendas.delete(0, tk.END)
        
        if self.produtos_queda_original is not None:
            self.produtos_queda = self.produtos_queda_original.copy()
            self.produtos_queda['Rank'] = range(1, len(self.produtos_queda) + 1)
            self.atualizar_lista()
    
    def atualizar_grafico(self):
        """Atualiza o gráfico de queda de vendas"""
        if self.produtos_queda is None or len(self.produtos_queda) == 0:
            return
        
        # Limpar gráfico anterior
        if self.canvas:
            self.canvas.get_tk_widget().destroy()
        
        # Criar novo gráfico (tamanho maior já que está em aba separada)
        self.fig, ax = plt.subplots(figsize=(10, 7))
        
        # Pegar top 10 para o gráfico
        top_10 = self.produtos_queda.head(10)
        
        # Criar gráfico de barras
        produtos = top_10['Produto'].values
        variacoes = top_10['Variação_%'].values
        
        # NÃO truncar nomes - usar nomes completos
        produtos_completos = list(produtos)
        
        cores = plt.cm.RdYlGn_r(np.linspace(0.2, 0.8, len(produtos)))
        bars = ax.barh(produtos_completos, variacoes, color=cores)
        
        ax.set_xlabel('Variação Percentual (%)', fontsize=10)
        ax.set_title('Top 10 Produtos com Maior Queda de Vendas', fontsize=12, fontweight='bold')
        ax.grid(axis='x', alpha=0.3)
        
        # Ajustar fonte dos nomes dos produtos para caber melhor
        ax.tick_params(axis='y', labelsize=8)
        
        # Adicionar valores nas barras
        for i, (bar, val) in enumerate(zip(bars, variacoes)):
            ax.text(val, i, f' {val:.1f}%', va='center', fontsize=8)
        
        # Ajustar layout para acomodar nomes completos
        plt.tight_layout()
        
        # Integrar com tkinter
        self.canvas = FigureCanvasTkAgg(self.fig, self.chart_frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
    
    def atualizar_grafico_pizza(self):
        """Atualiza o gráfico de pizza mostrando queda de vendas por linha"""
        if self.produtos_queda is None or len(self.produtos_queda) == 0:
            return
        
        # Limpar gráfico anterior
        if self.canvas_pizza:
            self.canvas_pizza.get_tk_widget().destroy()
        
        # Calcular queda de vendas agregada por linha
        # Calcular variação média ponderada por quantidade vendida
        def calcular_variacao_ponderada(group):
            if group['Quantidade_Vendidas'].sum() == 0:
                return 0
            return (group['Variação_%'] * group['Quantidade_Vendidas']).sum() / group['Quantidade_Vendidas'].sum()
        
        queda_por_linha = self.produtos_queda.groupby('Linha').agg({
            'Quantidade_Vendidas': 'sum',  # Soma total de quantidade vendida
            'Produto': 'count'  # Contagem de produtos
        }).reset_index()
        
        # Calcular variação média ponderada
        variacao_ponderada = self.produtos_queda.groupby('Linha').apply(calcular_variacao_ponderada).reset_index()
        variacao_ponderada.columns = ['Linha', 'Variacao_Media']
        
        # Merge com queda_por_linha
        queda_por_linha = queda_por_linha.merge(variacao_ponderada, on='Linha', how='left')
        queda_por_linha.columns = ['Linha', 'Qtd_Total', 'Num_Produtos', 'Variacao_Media']
        queda_por_linha['Variacao_Media'] = queda_por_linha['Variacao_Media'].fillna(0)
        
        # Calcular o impacto da queda: usar uma combinação de número de produtos e variação
        # Quanto mais negativo a variação e mais produtos, maior o impacto
        queda_por_linha['Impacto_Queda'] = abs(queda_por_linha['Variacao_Media']) * queda_por_linha['Num_Produtos']
        
        # Ordenar por impacto de queda (maior impacto primeiro)
        queda_por_linha = queda_por_linha.sort_values('Impacto_Queda', ascending=False)
        
        # Criar gráfico de pizza
        self.fig_pizza, ax = plt.subplots(figsize=(10, 7))
        
        # Preparar dados para o gráfico
        linhas = queda_por_linha['Linha'].values
        impactos = queda_por_linha['Impacto_Queda'].values
        
        # Se houver muitas linhas, mostrar apenas as top 10 e agrupar o resto como "Outros"
        if len(linhas) > 10:
            top_10 = impactos[:10]
            outros = impactos[10:].sum()
            linhas_plot = list(linhas[:10]) + ['Outros']
            impactos_plot = list(top_10) + [outros]
        else:
            linhas_plot = linhas
            impactos_plot = impactos
        
        # Criar gráfico de pizza
        cores = plt.cm.Set3(np.linspace(0, 1, len(linhas_plot)))
        wedges, texts, autotexts = ax.pie(
            impactos_plot, 
            labels=linhas_plot, 
            autopct='%1.1f%%',
            startangle=90,
            colors=cores,
            textprops={'fontsize': 9}
        )
        
        # Melhorar legibilidade dos textos
        for autotext in autotexts:
            autotext.set_color('black')
            autotext.set_fontweight('bold')
        
        ax.set_title('Distribuição de Queda de Vendas por Linha\n(Baseado no Impacto: Variação % × Número de Produtos)', 
                    fontsize=12, fontweight='bold', pad=20)
        
        # Adicionar legenda com informações adicionais
        legenda_info = []
        for idx, row in queda_por_linha.head(10).iterrows():
            legenda_info.append(
                f"{row['Linha']}: {row['Num_Produtos']:.0f} produtos, "
                f"Var. {row['Variacao_Media']:.1f}%"
            )
        
        # Criar texto de legenda
        legenda_texto = '\n'.join(legenda_info[:10])  # Top 10
        if len(queda_por_linha) > 10:
            legenda_texto += f"\n\n... e mais {len(queda_por_linha) - 10} linhas"
        
        # Adicionar texto informativo
        ax.text(0, -1.3, legenda_texto, 
               fontsize=8, ha='center', va='top',
               bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))
        
        plt.tight_layout()
        
        # Integrar com tkinter
        self.canvas_pizza = FigureCanvasTkAgg(self.fig_pizza, self.chart_pizza_frame)
        self.canvas_pizza.draw()
        self.canvas_pizza.get_tk_widget().grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
    
    def exportar_resultados(self):
        """Exporta os resultados para CSV"""
        if self.produtos_queda is None or len(self.produtos_queda) == 0:
            messagebox.showwarning("Aviso", "Nenhum resultado disponível para exportar")
            return
        
        arquivo = filedialog.asksaveasfilename(
            title="Salvar resultados",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if arquivo:
            try:
                # Preparar dados para exportação
                dados_export = self.produtos_queda.copy()
                
                # Formatar variação de preço para exportação
                if 'Variacao_Preco_%' in dados_export.columns:
                    dados_export['Variacao_Preco_%'] = dados_export['Variacao_Preco_%'].apply(
                        lambda x: f"{x:.1f}%" if pd.notna(x) and isinstance(x, (int, float)) else 'N/A'
                    )
                
                # Selecionar colunas para exportação
                colunas_export = ['Rank', 'Produto', 'Linha', 'Quantidade_Vendidas', 
                                'Mes_Inicio_Queda', 'Variacao_Preco_%', 
                                'Ultima_Venda', 'Mes_Mais_Vendas', 'Qtd_Mes_Mais_Vendas', 
                                'Mes_Menos_Vendas', 'Qtd_Ultimo_Mes', 'Media_Vendas']
                
                # Verificar quais colunas existem
                colunas_existentes = [col for col in colunas_export if col in dados_export.columns]
                dados_export = dados_export[colunas_existentes]
                
                # Formatar média de vendas
                if 'Media_Vendas' in dados_export.columns:
                    dados_export['Media_Vendas'] = dados_export['Media_Vendas'].apply(
                        lambda x: f"{x:.2f}" if pd.notna(x) else 'N/A'
                    )
                
                # Renomear colunas para português
                mapeamento_colunas = {
                    'Rank': 'Rank',
                    'Produto': 'Produto',
                    'Linha': 'Linha',
                    'Quantidade_Vendidas': 'Quantidade Vendidas',
                    'Mes_Inicio_Queda': 'Mês Início Queda',
                    'Variacao_Preco_%': 'Variação Preço %',
                    'Ultima_Venda': 'Última Venda',
                    'Mes_Mais_Vendas': 'Mês Mais Vendas',
                    'Qtd_Mes_Mais_Vendas': 'Quantidade Mês Mais Vendas',
                    'Mes_Menos_Vendas': 'Mês Menos Vendas',
                    'Qtd_Ultimo_Mes': 'Quantidade Último Mês',
                    'Mes_Menos_2': 'Mês -2',
                    'Qtd_Mes_Menos_2': 'Quantidade Mês -2',
                    'Mes_Menos_1': 'Mês -1',
                    'Qtd_Mes_Menos_1': 'Quantidade Mês -1',
                    'Mes_Atual': 'Mês Atual',
                    'Qtd_Mes_Atual': 'Quantidade Mês Atual',
                    'Media_Vendas': 'Média Vendas'
                }
                
                dados_export.columns = [mapeamento_colunas.get(col, col) for col in dados_export.columns]
                
                dados_export.to_csv(arquivo, index=False, encoding='utf-8-sig', sep=';', decimal=',')
                messagebox.showinfo("Sucesso", f"Resultados exportados com sucesso!\n{arquivo}")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao exportar resultados:\n{str(e)}")

def main():
    root = tk.Tk()
    app = AnalisadorVendas(root)
    root.mainloop()

if __name__ == "__main__":
    main()

