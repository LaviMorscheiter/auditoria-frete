import pandas as pd
import numpy as np
import re
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from dataclasses import dataclass
from functools import lru_cache
from typing import Dict, Tuple, List, Optional
import threading

try:
    sys.stdout.reconfigure(encoding='utf-8')
except:
    pass

# ================================================================
# CONFIGURAÇÕES
# ================================================================
MARGEM_TOLERANCIA = 15.00
LIMITE_PERCENTUAL_CRITICO = 0.10

HUB_CENTRAL = {
    "SAO PAULO","SÃO PAULO","BARUERI","SANTANA DE PARNAIBA","SANTANA DE PARNAÍBA",
    "OSASCO","GUARULHOS","CAJAMAR","COTIA","ITAPEVI","JANDIRA","CARAPICUIBA",
    "TABOAO DA SERRA","EMBU DAS ARTES","ITAQUAQUECETUBA","MAUA","MOGI DAS CRUZES",
    "SUZANO","SANTO ANDRE","SAO BERNARDO DO CAMPO","SAO CAETANO DO SUL","DIADEMA"
}

MAPA_UF_CAPITAL = {
    'AC':'RIO BRANCO','AL':'MACEIO','AM':'MANAUS','AP':'MACAPA',
    'BA':'SALVADOR','CE':'FORTALEZA','DF':'BRASILIA','ES':'VITORIA',
    'GO':'GOIANIA','MA':'SAO LUIS','MG':'BELO HORIZONTE','MS':'CAMPO GRANDE',
    'MT':'CUIABA','PA':'BELEM','PB':'JOAO PESSOA','PE':'RECIFE',
    'PI':'TERESINA','PR':'CURITIBA','RJ':'RIO DE JANEIRO','RN':'NATAL',
    'RO':'PORTO VELHO','RR':'BOA VISTA','RS':'PORTO ALEGRE','SC':'FLORIANOPOLIS',
    'SE':'ARACAJU','SP':'SAO PAULO','TO':'PALMAS'
}

TIPO_DIRETO = "DIRETO"
TIPO_SP_CAPITAL = "SP_CAPITAL"
TIPO_SP_INTERIOR = "SP_INTERIOR"
TIPO_BASE_CAPITAL = "BASE_CAPITAL"
TIPO_REDESPACHO = "REDESPACHO"

MAPA_ACENTOS = str.maketrans("ÁÀÃÂÄÉÈÊËÍÌÎÏÓÒÕÔÖÚÙÛÜÇ", "AAAAAEEEEIIIIOOOOOUUUUC")

# ================================================================
# FUNÇÕES AUXILIARES
# ================================================================
def limpar_texto(texto) -> str:
    if pd.isna(texto) or texto == "":
        return ""
    txt = str(texto).upper().strip().translate(MAPA_ACENTOS)
    return re.sub(r'\s{2,}', ' ', txt)

def quebrar_nome_coluna(col: str) -> List[str]:
    return [limpar_texto(p) for p in re.split(r'[,/ -]+', col)]

def formatar_moeda(valor):
    if pd.isna(valor):
        return "-"
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# ================================================================
# CLASSES
# ================================================================
@dataclass
class LPUContext:
    df_lpu: pd.DataFrame
    kg_adicional: Dict[str, float]
    col_redespacho: str

class AuditoriaFrete:
    def __init__(self, ctx: LPUContext):
        self.ctx = ctx
        self.colunas_lpu = list(ctx.df_lpu.columns)

    @lru_cache(maxsize=4096)
    def encontrar_coluna(self, cidade: str, uf: str) -> Tuple[Optional[str], str]:
        cidade_limpa = limpar_texto(cidade)
        
        if "BELO HOR" in cidade_limpa: cidade_limpa = "BELO HORIZONTE"
        if "R DE JANEIRO" in cidade_limpa: cidade_limpa = "RIO DE JANEIRO"

        for col in self.colunas_lpu:
            if cidade_limpa in quebrar_nome_coluna(col):
                return col, TIPO_DIRETO

        if uf == 'SP':
            tipo = TIPO_SP_CAPITAL if cidade_limpa in HUB_CENTRAL else TIPO_SP_INTERIOR
            for col in self.colunas_lpu:
                nome = limpar_texto(col)
                if ("SP" in nome) and (("CAPITAL" in nome and tipo == TIPO_SP_CAPITAL) or ("INTERIOR" in nome and tipo == TIPO_SP_INTERIOR)):
                    return col, tipo

        capital = MAPA_UF_CAPITAL.get(uf, "")
        if capital:
            eh_capital = (cidade_limpa in capital) or (capital in cidade_limpa)
            if eh_capital:
                for col in self.colunas_lpu:
                    if capital in limpar_texto(col):
                        return col, TIPO_BASE_CAPITAL

        return None, "NAO_ENCONTRADO"

    def calcular_valor(self, peso: float, coluna: str) -> float:
        if not coluna: return 0.0
        try:
            p_int = max(1, int(np.ceil(peso)))
            p_tab = min(p_int, 30)
            base = 0.0
            if p_tab in self.ctx.df_lpu.index:
                val = self.ctx.df_lpu.loc[p_tab, coluna]
                base = float(val) if pd.notna(val) else 0.0
            
            if p_int > 30:
                adic = float(self.ctx.kg_adicional.get(coluna, 0.0) or 0.0)
                return base + (p_int - 30) * adic
            return base
        except Exception:
            return 0.0

    def identificar_rota(self, cidade: str, uf: str) -> List[str]:
        col, tipo = self.encontrar_coluna(cidade, uf)
        cols_usadas = []
        
        if col:
            cols_usadas.append(col)
            capital_nome = MAPA_UF_CAPITAL.get(uf, "")
            cidade_limpa = limpar_texto(cidade)
            eh_a_propria_capital = (cidade_limpa in capital_nome) or (capital_nome in cidade_limpa)
            
            if tipo == TIPO_BASE_CAPITAL and not eh_a_propria_capital:
                cols_usadas.append(self.ctx.col_redespacho)
        else:
            capital_nome = MAPA_UF_CAPITAL.get(uf, "")
            col_capital = None
            if capital_nome:
                for c in self.colunas_lpu:
                    if capital_nome in limpar_texto(c):
                        col_capital = c
                        break
            
            if col_capital:
                cols_usadas.append(col_capital)
                cols_usadas.append(self.ctx.col_redespacho)
            else:
                cols_usadas.append(self.ctx.col_redespacho)
            
        return cols_usadas

    def investigar_peso_usado(self, valor_cobrado: float, cols_rota: List[str], peso_certo: float, vol: float) -> str:
        if valor_cobrado <= 0 or not cols_rota: return ""
        def simular(p): return sum(self.calcular_valor(p, c) for c in cols_rota)

        if vol > 0 and vol != peso_certo:
            if abs(simular(vol) - valor_cobrado) <= MARGEM_TOLERANCIA:
                return f"[!] ERRO: Cobrado por VOLUMES ({int(vol)}kg)"

        for p_test in range(1, 31):
            if abs(simular(p_test) - valor_cobrado) <= MARGEM_TOLERANCIA:
                return f"[!] ERRO: Cobrado como {p_test}kg (Tabela)"

        base30_total = sum(self.calcular_valor(30, c) for c in cols_rota)
        adic_total = sum(self.ctx.kg_adicional.get(c, 0.0) for c in cols_rota)
        
        if adic_total > 0 and valor_cobrado > base30_total:
            peso_calc = 30 + (valor_cobrado - base30_total) / adic_total
            if abs(peso_calc - round(peso_calc)) < 0.1:
                return f"[!] ERRO: Cobrado como {int(round(peso_calc))}kg (Cálc. Reverso)"
        return ""

    def auditar_linha(self, row: pd.Series) -> pd.Series:
        pr = pd.to_numeric(row.get("PESO REAL", 0), errors='coerce') or 0
        pc = pd.to_numeric(row.get("PESO CUBADO", 0), errors='coerce') or 0
        vol = pd.to_numeric(row.get("VOL", 0), errors='coerce') or 0
        peso_certo = max(pr, pc)

        cid_o = limpar_texto(row.get("CIDADE", ""))
        uf_o = limpar_texto(row.get("UF", ""))
        cid_d = limpar_texto(row.get("CIDADE.1", ""))
        uf_d = limpar_texto(row.get("UF.1", ""))
        frete_nf = pd.to_numeric(row.get("FRETE TOTAL", 0), errors='coerce') or 0

        cols_rota_total = []
        memoria = []
        custo_total = 0.0
        
        if cid_o and cid_o not in HUB_CENTRAL:
            cols = self.identificar_rota(cid_o, uf_o)
            cols_rota_total.extend(cols)
            custo_total += sum(self.calcular_valor(peso_certo, c) for c in cols)
            memoria.append(f"ORIG[{'+'.join(cols)}]")

        if cid_d and cid_d not in HUB_CENTRAL:
            cols = self.identificar_rota(cid_d, uf_d)
            cols_rota_total.extend(cols)
            custo_total += sum(self.calcular_valor(peso_certo, c) for c in cols)
            memoria.append(f"DEST[{'+'.join(cols)}]")

        if custo_total == 0:
            col_sp = next((c for c in self.colunas_lpu if "SP" in c and "CAPITAL" in c), self.colunas_lpu[0])
            cols_rota_total.append(col_sp)
            custo_total = self.calcular_valor(peso_certo, col_sp)
            memoria.append("SP_LOCAL")

        diff = frete_nf - custo_total
        percentual = 0.0
        if custo_total > 0:
            percentual = diff / custo_total

        status = "OK"
        sugestao = "-"
        
        if custo_total == 0:
            status = "ERRO_CALCULO"
        elif frete_nf == 0:
            status = "NF_ZERADA"
        elif abs(diff) <= MARGEM_TOLERANCIA:
            status = "OK"
        else:
            eh_critico = abs(percentual) > LIMITE_PERCENTUAL_CRITICO
            
            if diff > 0:
                if eh_critico:
                    status = "DIVERGENCIA_CRITICA"
                    sugestao = f"Pago a mais: {percentual:.0%} (Acima de 10%)"
                else:
                    status = "OK"
            else:
                investigacao = self.investigar_peso_usado(frete_nf, cols_rota_total, peso_certo, vol)
                if investigacao:
                    status = "ERRO_PESO_INCORRETO"
                    sugestao = investigacao
                elif eh_critico:
                    status = "DIVERGENCIA_CRITICA"
                    sugestao = f"Pago a menos: {abs(percentual):.0%} (Acima de 10%)"
                else:
                    status = "OK"

        return pd.Series([
            int(np.ceil(peso_certo)),
            round(custo_total, 2),
            round(diff, 2),
            status,
            " + ".join(memoria),
            sugestao
        ])

# ================================================================
# INTERFACE GRÁFICA
# ================================================================
class AuditoriaFreteGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Auditoria de Frete - Sistema Profissional")
        self.root.geometry("750x700")
        self.root.resizable(True, True)
        
        # Variáveis
        self.lpu_path = tk.StringVar()
        self.rel_path = tk.StringVar()
        self.resultado = None
        
        # Estilo
        style = ttk.Style()
        style.theme_use('clam')
        
        self.criar_interface()
        
    def criar_interface(self):
        # Frame Principal
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        title_label = tk.Label(
            main_frame, 
            text="🚚 AUDITORIA DE FRETE", 
            font=("Arial", 22, "bold"),
            fg="#1e3a8a"
        )
        title_label.pack(pady=(0, 15))
        
        # Seção Upload LPU
        lpu_frame = ttk.LabelFrame(main_frame, text="1. Tabela LPU", padding="10")
        lpu_frame.pack(fill=tk.X, pady=5)
        
        ttk.Entry(lpu_frame, textvariable=self.lpu_path, state="readonly", width=65).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(lpu_frame, text="📁 Selecionar", command=self.selecionar_lpu).pack(side=tk.LEFT)
        
        # Seção Upload Relatório
        rel_frame = ttk.LabelFrame(main_frame, text="2. Relatório de Fretes", padding="10")
        rel_frame.pack(fill=tk.X, pady=5)
        
        ttk.Entry(rel_frame, textvariable=self.rel_path, state="readonly", width=65).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(rel_frame, text="📁 Selecionar", command=self.selecionar_relatorio).pack(side=tk.LEFT)
        
        # Botão Processar
        self.btn_processar = tk.Button(
            main_frame,
            text="⚙️ AUDITAR FRETES",
            font=("Arial", 13, "bold"),
            bg="#2563eb",
            fg="white",
            pady=12,
            cursor="hand2",
            command=self.processar_auditoria,
            relief=tk.RAISED,
            borderwidth=2
        )
        self.btn_processar.pack(fill=tk.X, pady=15)
        
        # Frame Resultados
        self.result_frame = ttk.LabelFrame(main_frame, text="📊 Resultados da Auditoria", padding="15")
        self.result_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Labels de Resultados com DESTAQUE
        self.label_pago = tk.Label(
            self.result_frame, 
            text="💰 Total Pago pela Empresa: Aguardando...", 
            font=("Arial", 12, "bold"), 
            anchor="w",
            bg="#dbeafe",
            fg="#1e40af",
            padx=12,
            pady=10,
            relief=tk.SOLID,
            borderwidth=1
        )
        self.label_pago.pack(fill=tk.X, pady=5)
        
        self.label_lpu = tk.Label(
            self.result_frame, 
            text="📋 Valor Calculado LPU: Aguardando...", 
            font=("Arial", 12, "bold"), 
            anchor="w",
            bg="#dbeafe",
            fg="#1e40af",
            padx=12,
            pady=10,
            relief=tk.SOLID,
            borderwidth=1
        )
        self.label_lpu.pack(fill=tk.X, pady=5)
        
        self.label_diff = tk.Label(
            self.result_frame, 
            text="📊 Diferença: Aguardando...", 
            font=("Arial", 13, "bold"), 
            anchor="w",
            bg="#fef3c7",
            fg="#92400e",
            padx=12,
            pady=12,
            relief=tk.SOLID,
            borderwidth=2
        )
        self.label_diff.pack(fill=tk.X, pady=5)
        
        # Separador visual FORTE
        separator = tk.Frame(self.result_frame, height=3, bg="#9ca3af")
        separator.pack(fill=tk.X, pady=15)
        
        # Label indicativo
        label_download = tk.Label(
            self.result_frame,
            text="👇 Clique no botão abaixo para baixar o relatório completo",
            font=("Arial", 10, "italic"),
            fg="#6b7280"
        )
        label_download.pack(pady=(0, 10))
        
        # Botão Download GRANDE E VISÍVEL
        self.btn_download = tk.Button(
            self.result_frame,
            text="💾 BAIXAR RELATÓRIO COMPLETO (EXCEL)",
            font=("Arial", 14, "bold"),
            bg="#16a34a",
            fg="white",
            activebackground="#15803d",
            activeforeground="white",
            pady=20,
            state=tk.DISABLED,
            cursor="hand2",
            command=self.baixar_relatorio,
            relief=tk.RAISED,
            borderwidth=4
        )
        self.btn_download.pack(fill=tk.X, pady=(0, 10), ipady=8)
        
        # Status
        self.status_label = tk.Label(main_frame, text="Aguardando arquivos...", font=("Arial", 10), fg="#6b7280")
        self.status_label.pack(pady=(10, 0))
    
    def selecionar_lpu(self):
        arquivo = filedialog.askopenfilename(
            title="Selecione a Tabela LPU",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("Todos", "*.*")]
        )
        if arquivo:
            self.lpu_path.set(arquivo)
            self.status_label.config(text=f"✓ LPU selecionada: {os.path.basename(arquivo)}", fg="#16a34a")
    
    def selecionar_relatorio(self):
        arquivo = filedialog.askopenfilename(
            title="Selecione o Relatório de Fretes",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("Todos", "*.*")]
        )
        if arquivo:
            self.rel_path.set(arquivo)
            self.status_label.config(text=f"✓ Relatório selecionado: {os.path.basename(arquivo)}", fg="#16a34a")
    
    def processar_auditoria(self):
        if not self.lpu_path.get() or not self.rel_path.get():
            messagebox.showwarning("Atenção", "Por favor, selecione ambos os arquivos!")
            return
        
        self.btn_processar.config(state=tk.DISABLED, text="⏳ Processando...")
        self.status_label.config(text="⏳ Processando auditoria... Aguarde, isto pode levar alguns segundos.", fg="#ea580c")
        self.root.update()
        
        # Processar em thread
        thread = threading.Thread(target=self._processar_thread)
        thread.start()
    
    def _processar_thread(self):
        try:
            # Carregar LPU
            path_lpu = self.lpu_path.get()
            if path_lpu.endswith('.csv'):
                with open(path_lpu, 'r', encoding='latin1') as f: 
                    lines = f.readlines()
                idx = next((i for i, l in enumerate(lines[:20]) if "Peso" in l), 0)
                df_lpu = pd.read_csv(path_lpu, header=idx)
            else:
                df_lpu = pd.read_excel(path_lpu, header=8)

            if "Peso" in df_lpu.columns: 
                df_lpu.set_index("Peso", inplace=True)
            
            df_lpu.columns = df_lpu.columns.str.replace(r'\s{2,}', ' ', regex=True).str.strip().str.upper()
            df_lpu = df_lpu.loc[:, ~df_lpu.columns.str.contains('^UNNAMED')]

            idx_30 = df_lpu.index[pd.to_numeric(df_lpu.index, errors='coerce') == 30]
            if len(idx_30) == 0: 
                raise ValueError("Linha 30kg não encontrada na LPU")
            
            pos_30 = df_lpu.index.get_loc(idx_30[0])
            linha_adic = df_lpu.iloc[pos_30 + 1]
            kg_adicional = {col: float(linha_adic[col]) for col in df_lpu.columns}
            df_lpu = df_lpu.iloc[:pos_30 + 1]
            df_lpu.index = pd.to_numeric(df_lpu.index, errors='coerce')
            
            col_red = next((c for c in df_lpu.columns[::-1] if "REDESPACHO" in c or "INTERIOR" in c), df_lpu.columns[-1])
            ctx = LPUContext(df_lpu=df_lpu, kg_adicional=kg_adicional, col_redespacho=col_red)
            auditor = AuditoriaFrete(ctx)

            # Carregar Relatório
            path_rel = self.rel_path.get()
            if path_rel.endswith('.csv'):
                df_rel = pd.read_csv(path_rel, sep=None, engine='python')
            else:
                df_rel = pd.read_excel(path_rel)
            
            df_rel.columns = df_rel.columns.str.upper().str.strip()

            # Auditar
            resultado = df_rel.apply(auditor.auditar_linha, axis=1)
            resultado.columns = ['PESO_COBRAVEL', 'VALOR_LPU', 'DIFERENCA', 'STATUS', 'MEMORIA_TECNICA', 'SUGESTAO']
            df_final = pd.concat([df_rel, resultado], axis=1)

            # Preparar Export
            mapa_colunas = {
                "PESO REAL": "PESO REAL", "PESO CUBADO": "PESO CUBADO", "VOL": "VOLUMES",
                "CIDADE": "ORIGEM_CIDADE", "UF": "ORIGEM_UF",
                "CIDADE.1": "DESTINO_CIDADE", "UF.1": "DESTINO_UF",
                "FRETE TOTAL": "FRETE TOTAL", "PESO_COBRAVEL": "PESO_COBRAVEL",
                "VALOR_LPU": "VALOR_LPU", "DIFERENCA": "DIFERENCA",
                "STATUS": "STATUS", "MEMORIA_TECNICA": "MEMORIA_TECNICA", "SUGESTAO": "SUGESTAO"
            }

            df_export = pd.DataFrame()
            for col_orig, col_dest in mapa_colunas.items():
                if col_orig in df_final.columns:
                    df_export[col_dest] = df_final[col_orig]

            # Calcular Totais
            soma_gasto = df_export['FRETE TOTAL'].sum()
            soma_devido = df_export['VALOR_LPU'].sum()
            soma_diferenca = df_export['DIFERENCA'].sum()

            # Adicionar linha de totais
            new_row = {col: np.nan for col in df_export.columns}
            new_row['ORIGEM_CIDADE'] = 'TOTAL GERAL'
            new_row['FRETE TOTAL'] = soma_gasto
            new_row['VALOR_LPU'] = soma_devido
            new_row['DIFERENCA'] = soma_diferenca
            new_row['DESTINO_UF'] = 'CONFIRA: (Gasto | Devido | Diferenca) ->'
            new_row['STATUS'] = ''
            new_row['SUGESTAO'] = ''
            
            df_export = pd.concat([df_export, pd.DataFrame([new_row])], ignore_index=True)

            # Estilo
            def highlight_critical(row):
                styles = [''] * len(row)
                if row.get('ORIGEM_CIDADE') == 'TOTAL GERAL':
                    return ['background-color: #D3D3D3; font-weight: bold; border-top: 2px solid black'] * len(row)
                status = str(row.get('STATUS', ''))
                if status == 'DIVERGENCIA_CRITICA':
                    return ['background-color: #FF5733; color: white; font-weight: bold'] * len(row)
                if status == 'ERRO_PESO_INCORRETO':
                    return ['background-color: #FFC300; font-weight: bold'] * len(row)
                return styles

            styled_df = df_export.style.apply(highlight_critical, axis=1)
            styled_df = styled_df.format({
                'FRETE TOTAL': formatar_moeda, 
                'VALOR_LPU': formatar_moeda, 
                'DIFERENCA': formatar_moeda
            }, na_rep="-")

            # Salvar resultado
            self.resultado = (styled_df, soma_gasto, soma_devido, soma_diferenca)
            
            # Atualizar interface na thread principal
            self.root.after(0, self._atualizar_resultados)
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Erro", f"Erro ao processar:\n\n{str(e)}"))
            self.root.after(0, self._reset_botoes)
    
    def _atualizar_resultados(self):
        if self.resultado:
            styled_df, soma_gasto, soma_devido, soma_diferenca = self.resultado
            
            # Atualizar labels com valores formatados
            self.label_pago.config(
                text=f"💰 Total Pago pela Empresa: {formatar_moeda(soma_gasto)}",
                bg="#dbeafe",
                fg="#1e40af"
            )
            
            self.label_lpu.config(
                text=f"📋 Valor Calculado LPU (Devido): {formatar_moeda(soma_devido)}",
                bg="#dbeafe",
                fg="#1e40af"
            )
            
            # Diferença com cor dinâmica
            if soma_diferenca > 0:
                cor_fundo = "#fecaca"
                cor_texto = "#991b1b"
                texto_tipo = "PREJUÍZO"
            else:
                cor_fundo = "#bbf7d0"
                cor_texto = "#166534"
                texto_tipo = "ECONOMIA"
            
            sinal = "+" if soma_diferenca > 0 else ""
            
            self.label_diff.config(
                text=f"📊 Diferença ({texto_tipo}): {sinal}{formatar_moeda(soma_diferenca)}", 
                bg=cor_fundo,
                fg=cor_texto
            )
            
            # Habilitar botão de download
            self.btn_download.config(state=tk.NORMAL)
            self.status_label.config(
                text="✅ Auditoria concluída! Clique no botão verde para baixar o relatório completo.", 
                fg="#16a34a"
            )
        
        self._reset_botoes()
    
    def _reset_botoes(self):
        self.btn_processar.config(state=tk.NORMAL, text="⚙️ AUDITAR FRETES")
    
    def baixar_relatorio(self):
        if not self.resultado:
            messagebox.showwarning("Atenção", "Não há resultado para baixar!")
            return
        
        arquivo = filedialog.asksaveasfilename(
            title="Salvar Relatório de Auditoria",
            defaultextension=".xlsx",
            initialfile="Auditoria_Frete_Completa.xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        
        if arquivo:
            try:
                styled_df, _, _, _ = self.resultado
                styled_df.to_excel(arquivo, index=False, engine='openpyxl')
                self.status_label.config(
                    text=f"✅ Arquivo salvo com sucesso: {os.path.basename(arquivo)}", 
                    fg="#16a34a"
                )
                messagebox.showinfo(
                    "Sucesso!", 
                    f"Relatório salvo com sucesso!\n\nLocal: {arquivo}\n\nO arquivo Excel contém:\n• Todas as linhas auditadas\n• Cores para divergências críticas\n• Linha de totais no final"
                )
                
                # Tentar abrir a pasta
                try:
                    os.startfile(os.path.dirname(arquivo))
                except:
                    pass
                    
            except Exception as e:
                messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar o arquivo:\n\n{str(e)}")

# ================================================================
# MAIN
# ================================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = AuditoriaFreteGUI(root)
    root.mainloop()