import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import threading
import os
from datetime import datetime

class CSVtoExcelConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("🎬 Conversor CSV → Excel")
        self.root.geometry("800x600")
        self.root.configure(bg="#1a1a2e")
        
        # Variáveis
        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.max_rows = tk.IntVar(value=200000)
        self.sample_size = tk.IntVar(value=500)
        self.separator = tk.StringVar(value="\t")
        self.encoding = tk.StringVar(value="utf-8")
        self.is_processing = False
        
        self.setup_ui()
        
    def setup_ui(self):
        # Estilo moderno
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TFrame", background="#1a1a2e")
        style.configure("TLabel", background="#1a1a2e", foreground="#eee", font=('Segoe UI', 10))
        style.configure("TButton", font=('Segoe UI', 10, 'bold'), padding=10)
        
        # Container principal
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        title = tk.Label(main_frame, text="📊 Conversor CSV para Excel", 
                        font=('Segoe UI', 20, 'bold'), bg="#1a1a2e", fg="#e94560")
        title.pack(pady=(0, 20))
        
        # Área de arquivo de entrada
        input_frame = ttk.LabelFrame(main_frame, text=" Arquivo de Entrada (CSV) ", padding="10")
        input_frame.pack(fill=tk.X, pady=5)
        
        ttk.Entry(input_frame, textvariable=self.input_file, font=('Consolas', 10)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        ttk.Button(input_frame, text="📁 Procurar", command=self.browse_input).pack(side=tk.RIGHT)
        
        # Opções de configuração
        options_frame = ttk.LabelFrame(main_frame, text=" Configurações ", padding="10")
        options_frame.pack(fill=tk.X, pady=10)
        
        # Grid de opções
        ttk.Label(options_frame, text="Separador:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        sep_combo = ttk.Combobox(options_frame, textvariable=self.separator, values=["\t", ",", ";", "|"], width=10, state="readonly")
        sep_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        sep_combo.set("\t")
        
        ttk.Label(options_frame, text="Encoding:").grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
        ttk.Combobox(options_frame, textvariable=self.encoding, values=["utf-8", "latin-1", "iso-8859-1", "cp1252"], width=12, state="readonly").grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(options_frame, text="Máx. Linhas:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Spinbox(options_frame, from_=1000, to=1000000, textvariable=self.max_rows, width=12).grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(options_frame, text="Tamanho Amostra:").grid(row=1, column=2, sticky=tk.W, padx=5, pady=5)
        ttk.Spinbox(options_frame, from_=100, to=10000, textvariable=self.sample_size, width=12).grid(row=1, column=3, sticky=tk.W, padx=5, pady=5)
        
        # Área de arquivo de saída
        output_frame = ttk.LabelFrame(main_frame, text=" Arquivo de Saída (Excel) ", padding="10")
        output_frame.pack(fill=tk.X, pady=5)
        
        ttk.Entry(output_frame, textvariable=self.output_file, font=('Consolas', 10)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        ttk.Button(output_frame, text="💾 Guardar Como", command=self.browse_output).pack(side=tk.RIGHT)
        
        # Barra de progresso
        self.progress = ttk.Progressbar(main_frame, mode='determinate', length=400)
        self.progress.pack(pady=15)
        
        # Botão de conversão
        self.convert_btn = tk.Button(main_frame, text="🚀 Converter para Excel", 
                                    font=('Segoe UI', 12, 'bold'), bg="#e94560", fg="white",
                                    activebackground="#ff6b6b", activeforeground="white",
                                    relief=tk.FLAT, cursor="hand2", command=self.start_conversion)
        self.convert_btn.pack(pady=10, ipadx=20, ipady=10)
        
        # Log de atividades
        log_frame = ttk.LabelFrame(main_frame, text=" Log de Processamento ", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=8, font=('Consolas', 9), 
                                                 bg="#16213e", fg="#00ff88", insertbackground="white")
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Status bar
        self.status_var = tk.StringVar(value="Pronto para converter")
        status_bar = tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, 
                             anchor=tk.W, bg="#0f3460", fg="white", font=('Segoe UI', 9))
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def browse_input(self):
        filename = filedialog.askopenfilename(
            title="Selecionar arquivo CSV",
            filetypes=[("CSV files", "*.csv *.tsv *.txt"), ("All files", "*.*")]
        )
        if filename:
            self.input_file.set(filename)
            if not self.output_file.get():
                base = os.path.splitext(filename)[0]
                self.output_file.set(f"{base}.xlsx")
            self.log(f"Arquivo selecionado: {os.path.basename(filename)}")
            
    def browse_output(self):
        filename = filedialog.asksaveasfilename(
            title="Guardar arquivo Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.output_file.set(filename)
            self.log(f"Destino definido: {os.path.basename(filename)}")
            
    def start_conversion(self):
        if self.is_processing:
            return
            
        if not self.input_file.get():
            messagebox.showerror("Erro", "Por favor, selecione um arquivo CSV!")
            return
            
        if not self.output_file.get():
            messagebox.showerror("Erro", "Por favor, defina o arquivo de saída!")
            return
            
        self.is_processing = True
        self.convert_btn.config(state=tk.DISABLED, text="⏳ Processando...")
        self.progress.start()
        
        thread = threading.Thread(target=self.convert_file)
        thread.daemon = True
        thread.start()
        
    def convert_file(self):
        try:
            caminho = self.input_file.get()
            saida = self.output_file.get()
            max_linhas = self.max_rows.get()
            amostra = self.sample_size.get()
            sep = self.separator.get()
            enc = self.encoding.get()
            
            self.log("Iniciando leitura do CSV...")
            self.status_var.set("A ler arquivo CSV...")
            
            # Ler CSV
            df = pd.read_csv(caminho, sep=sep, nrows=max_linhas, dtype=str, encoding=enc)
            self.log(f"Carregadas {len(df)} linhas do CSV")
            
            # Limpar dados (baseado no seu código original)
            if 'titleType' in df.columns and 'primaryTitle' in df.columns:
                df = df[df['titleType'].notna() & df['primaryTitle'].notna()]
                self.log(f"Filtradas {len(df)} linhas válidas")
                
                if 'startYear' in df.columns:
                    df['startYear'] = pd.to_numeric(df['startYear'], errors='coerce')
                if 'runtimeMinutes' in df.columns:
                    df['runtimeMinutes'] = pd.to_numeric(df['runtimeMinutes'], errors='coerce')
                if 'genres' in df.columns:
                    df['genres'] = df['genres'].fillna('').apply(
                        lambda g: g.split(',') if g != '\\N' and str(g).strip() != '' else []
                    )
            
            # Amostragem
            if len(df) > amostra:
                df = df.sample(n=amostra, random_state=42).reset_index(drop=True)
                self.log(f"Amostra de {amostra} registros selecionada")
            
            self.status_var.set("A converter para Excel...")
            self.log("A criar arquivo Excel...")
            
            # Salvar Excel com formatação
            with pd.ExcelWriter(saida, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Dados', index=False)
                
                # Ajustar largura das colunas
                worksheet = writer.sheets['Dados']
                for idx, col in enumerate(df.columns):
                    max_length = max(
                        df[col].astype(str).map(len).max(),
                        len(str(col))
                    ) + 2
                    worksheet.column_dimensions[chr(65 + idx)].width = min(max_length, 50)
            
            self.log(f"✅ Conversão concluída!")
            self.log(f"📁 Arquivo guardado: {saida}")
            self.status_var.set(f"Concluído! {len(df)} registros exportados")
            
            stats = f"""
Estatísticas:
• Total de registros: {len(df)}
• Colunas: {len(df.columns)}
• Tamanho do arquivo: {os.path.getsize(saida) / 1024:.1f} KB
            """
            self.log(stats)
            
            messagebox.showinfo("Sucesso", f"Conversão concluída!\n\n{len(df)} registros exportados para:\n{saida}")
            
        except FileNotFoundError:
            self.log("❌ Erro: Ficheiro não encontrado")
            messagebox.showerror("Erro", "Ficheiro não encontrado!")
        except Exception as e:
            self.log(f"❌ Erro: {str(e)}")
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{str(e)}")
        finally:
            self.is_processing = False
            self.convert_btn.config(state=tk.NORMAL, text="🚀 Converter para Excel")
            self.progress.stop()
            self.progress['value'] = 0

# Iniciar aplicação
if __name__ == "__main__":
    root = tk.Tk()
    app = CSVtoExcelConverter(root)
    root.mainloop()