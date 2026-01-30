import pandas as pd
import numpy as np
import re
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from dataclasses import dataclass
from typing import Dict, Optional
import threading
import warnings

warnings.simplefilter("ignore")

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
    'AC': 'RIO BRANCO', 'AL': 'MACEIO', 'AM': 'MANAUS', 'AP': 'MACAPA',
    'BA': 'SALVADOR', 'CE': 'FORTALEZA', 'DF': 'BRASILIA', 'ES': 'VITORIA',
    'GO': 'GOIANIA', 'MA': 'SAO LUIS', 'MG': 'BELO HORIZONTE', 'MS': 'CAMPO GRANDE',
    'MT': 'CUIABA', 'PA': 'BELEM', 'PB': 'JOAO PESSOA', 'PE': 'RECIFE',
    'PI': 'TERESINA', 'PR': 'CURITIBA', 'RJ': 'RIO DE JANEIRO', 'RN': 'NATAL',
    'RO': 'PORTO VELHO', 'RR': 'BOA VISTA', 'RS': 'PORTO ALEGRE', 'SC': 'FLORIANOPOLIS',
    'SE': 'ARACAJU', 'SP': 'SAO PAULO', 'TO': 'PALMAS'
}

MAPA_ACENTOS = str.maketrans("ÁÀÃÂÄÉÈÊËÍÌÎÏÓÒÕÔÖÚÙÛÜÇ", "AAAAAEEEEIIIIOOOOOUUUUC")

# ================================================================
# FUNÇÕES UTILITÁRIAS
# ================================================================
def safe_float(valor):
    if isinstance(valor, pd.Series):
        valor = valor.iloc[0] if not valor.empty else 0.0
    if pd.isna(valor) or str(valor).strip() == "": 
        return 0.0
    if isinstance(valor, (int, float)): 
        return float(valor)
    try:
        v_str = str(valor).upper().replace("R$", "").strip()
        if "." in v_str and "," in v_str:
            v_str = v_str.replace(".", "")
        v_str = v_str.replace(",", ".")
        return float(v_str)
    except: 
        return 0.0

def limpar_texto(texto) -> str:
    if isinstance(texto, pd.Series):
        texto = texto.iloc[0] if not texto.empty else ""
    if pd.isna(texto) or texto == "": 
        return ""
    txt = str(texto).upper().strip().translate(MAPA_ACENTOS)
    return re.sub(r'\s{2,}', ' ', txt)

def formatar_moeda(valor):
    if pd.isna(valor): 
        return "-"
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# ================================================================
# LEITOR INTELIGENTE DE ARQUIVOS
# ================================================================

class LeitorArquivo:
    """Lê arquivos Excel/CSV de forma inteligente."""
    
    @staticmethod
    def carregar(caminho: str) -> pd.DataFrame:
        """Carrega arquivo detectando engine e formato."""
        ext = caminho.lower().split('.')[-1]
        
        if ext in ['xls', 'xlsx', 'xlsm']:
            engine = 'xlrd' if ext == 'xls' else 'openpyxl'
            return pd.read_excel(caminho, header=None, engine=engine)
        elif ext == 'csv':
            try:
                return pd.read_csv(caminho, header=None, sep=None, engine='python', encoding='latin1')
            except:
                return pd.read_csv(caminho, header=None, sep=';', encoding='utf-8')
        else:
            raise Exception(f"Formato não suportado: {ext}")
    
    @staticmethod
    def encontrar_cabecalho(df: pd.DataFrame, palavras_chave: list) -> int:
        """Encontra linha do cabeçalho baseado em palavras-chave."""
        for i in range(min(30, len(df))):
            row_str = " ".join([str(x).upper() for x in df.iloc[i].values])
            score = sum(1 for palavra in palavras_chave if palavra in row_str)
            if score >= 3:  # Precisa de pelo menos 3 matches
                return i
        return 0
    
    @staticmethod
    def deduplica_colunas(df: pd.DataFrame) -> pd.DataFrame:
        """Remove colunas duplicadas adicionando sufixos."""
        cols = []
        seen = {}
        for c in df.columns:
            c_str = str(c).upper().strip()
            if 'UNNAMED' in c_str:
                continue
            if c_str in seen:
                seen[c_str] += 1
                cols.append(f"{c_str}_DUP{seen[c_str]}")
            else:
                seen[c_str] = 0
                cols.append(c_str)
        
        df = df.iloc[:, :len(cols)]
        df.columns = cols
        return df.loc[:, ~df.columns.str.startswith('UNNAMED')]

# ================================================================
# EXTRATOR DE LOCALIZAÇÃO
# ================================================================

class ExtratorLocalizacao:
    """Extrai cidade e UF de diferentes formatos de colunas."""
    
    UFS_VALIDAS = set(MAPA_UF_CAPITAL.keys())
    
    @staticmethod
    def extrair_cidade_uf(texto: str) -> tuple:
        """
        Extrai cidade e UF de texto. Exemplos:
        - "SANTANA DE PARNAIBA - SP" → ("SANTANA DE PARNAIBA", "SP")
        - "SAO PAULO/SP" → ("SAO PAULO", "SP")
        - "CURITIBA" → ("CURITIBA", "")
        """
        texto_limpo = limpar_texto(texto)
        
        # Padrão: CIDADE - UF ou CIDADE/UF ou CIDADE (UF)
        padrao = r'(.+?)[\s\-/\(]*([A-Z]{2})[\s\)]*$'
        match = re.search(padrao, texto_limpo)
        
        if match:
            cidade = match.group(1).strip()
            uf = match.group(2).strip()
            
            # Valida se é realmente uma UF
            if uf in ExtratorLocalizacao.UFS_VALIDAS:
                return (cidade, uf)
        
        # Se não encontrou UF no texto, retorna só a cidade
        return (texto_limpo.strip(), "")
    
    @staticmethod
    def processar_linha(row: pd.Series, colunas_detectadas: dict) -> dict:
        """
        Processa uma linha e extrai origem/destino considerando formato da planilha.
        
        Retorna: {
            'origem_cidade': str,
            'origem_uf': str,
            'destino_cidade': str,
            'destino_uf': str
        }
        """
        resultado = {
            'origem_cidade': '',
            'origem_uf': '',
            'destino_cidade': '',
            'destino_uf': ''
        }
        
        # ORIGEM
        if 'origem_cidade' in colunas_detectadas:
            cidade_col = colunas_detectadas['origem_cidade']
            cidade_texto = str(row.get(cidade_col, ''))
            
            # Se tem coluna UF separada, usa ela
            if 'origem_uf' in colunas_detectadas:
                resultado['origem_cidade'] = limpar_texto(cidade_texto)
                resultado['origem_uf'] = limpar_texto(row.get(colunas_detectadas['origem_uf'], ''))
            else:
                # Extrai cidade e UF do mesmo campo
                cidade, uf = ExtratorLocalizacao.extrair_cidade_uf(cidade_texto)
                resultado['origem_cidade'] = cidade
                resultado['origem_uf'] = uf
        
        # DESTINO
        if 'destino_cidade' in colunas_detectadas:
            cidade_col = colunas_detectadas['destino_cidade']
            cidade_texto = str(row.get(cidade_col, ''))
            
            # Se tem coluna UF separada, usa ela
            if 'destino_uf' in colunas_detectadas:
                resultado['destino_cidade'] = limpar_texto(cidade_texto)
                resultado['destino_uf'] = limpar_texto(row.get(colunas_detectadas['destino_uf'], ''))
            else:
                # Extrai cidade e UF do mesmo campo
                cidade, uf = ExtratorLocalizacao.extrair_cidade_uf(cidade_texto)
                resultado['destino_cidade'] = cidade
                resultado['destino_uf'] = uf
        
        return resultado

# ================================================================
# DETECTOR DE ESTRUTURA
# ================================================================

class DetectorEstrutura:
    """Detecta automaticamente a estrutura das colunas da planilha."""
    
    @staticmethod
    def detectar(colunas: list) -> dict:
        """
        Detecta quais colunas existem e como estão organizadas.
        
        Retorna dict com mapeamento: {
            'origem_cidade': nome_coluna_real,
            'origem_uf': nome_coluna_real ou None,
            'destino_cidade': nome_coluna_real,
            'destino_uf': nome_coluna_real ou None,
            ...
        }
        """
        mapa = {}
        cols_usadas = set()
        
        colunas_limpas = [(i, c, limpar_texto(c)) for i, c in enumerate(colunas)]
        
        # Detecta PESOS
        for campo, palavras in [
            ('peso_real', ['PESO REAL', 'PESO', 'KG']),
            ('peso_cubado', ['PESO CUBADO', 'CUBADO', 'PESO CUB']),
            ('peso_taxado', ['PESO TAXADO', 'TAXADO', 'P. TAXADO']),
            ('frete_total', ['FRETE TOTAL', 'VALOR FRETE', 'VALOR TOTAL', 'TOTAL'])
        ]:
            col = DetectorEstrutura._buscar_coluna(colunas_limpas, palavras, cols_usadas)
            if col:
                mapa[campo] = col
                cols_usadas.add(col)
        
        # Detecta ORIGEM (pode ser REMETENTE ou ORIGEM)
        palavras_origem = ['REMETENTE', 'ORIGEM']
        col_origem = DetectorEstrutura._buscar_coluna(colunas_limpas, palavras_origem, cols_usadas)
        
        if col_origem:
            idx_origem = next(i for i, c, _ in colunas_limpas if c == col_origem)
            cols_usadas.add(col_origem)
            
            # Procura CIDADE depois de ORIGEM/REMETENTE
            col_cidade_origem = DetectorEstrutura._buscar_proxima(colunas_limpas, idx_origem, ['CIDADE'], cols_usadas)
            if col_cidade_origem:
                mapa['origem_cidade'] = col_cidade_origem
                cols_usadas.add(col_cidade_origem)
                
                # Procura UF depois da CIDADE
                col_uf_origem = DetectorEstrutura._buscar_proxima(colunas_limpas, 
                    next(i for i, c, _ in colunas_limpas if c == col_cidade_origem),
                    ['UF', 'ESTADO'], cols_usadas)
                if col_uf_origem:
                    mapa['origem_uf'] = col_uf_origem
                    cols_usadas.add(col_uf_origem)
        
        # Detecta DESTINO (pode ser DESTINATARIO ou DESTINO)
        palavras_destino = ['DESTINATARIO', 'DESTINATÁRIO', 'DESTINO']
        col_destino = DetectorEstrutura._buscar_coluna(colunas_limpas, palavras_destino, cols_usadas)
        
        if col_destino:
            idx_destino = next(i for i, c, _ in colunas_limpas if c == col_destino)
            cols_usadas.add(col_destino)
            
            # Procura CIDADE depois de DESTINO/DESTINATÁRIO
            col_cidade_destino = DetectorEstrutura._buscar_proxima(colunas_limpas, idx_destino, ['CIDADE'], cols_usadas)
            if col_cidade_destino:
                mapa['destino_cidade'] = col_cidade_destino
                cols_usadas.add(col_cidade_destino)
                
                # Procura UF depois da CIDADE
                col_uf_destino = DetectorEstrutura._buscar_proxima(colunas_limpas,
                    next(i for i, c, _ in colunas_limpas if c == col_cidade_destino),
                    ['UF', 'ESTADO'], cols_usadas)
                if col_uf_destino:
                    mapa['destino_uf'] = col_uf_destino
                    cols_usadas.add(col_uf_destino)
        
        return mapa
    
    @staticmethod
    def _buscar_coluna(colunas_limpas, palavras, ignorar):
        """Busca coluna que contém alguma das palavras."""
        for _, col_original, col_limpo in colunas_limpas:
            if col_original in ignorar:
                continue
            for palavra in palavras:
                if palavra in col_limpo:
                    return col_original
        return None
    
    @staticmethod
    def _buscar_proxima(colunas_limpas, idx_inicio, palavras, ignorar):
        """Busca próxima coluna após idx_inicio que contém alguma palavra."""
        for i, col_original, col_limpo in colunas_limpas:
            if i <= idx_inicio or col_original in ignorar:
                continue
            for palavra in palavras:
                if palavra in col_limpo:
                    return col_original
        return None

# ================================================================
# CALCULADORA DE PESO
# ================================================================

class CalculadoraPeso:
    """Determina peso correto e peso cobrado."""
    
    @staticmethod
    def processar(peso_real, peso_cubado, peso_taxado) -> tuple:
        """
        Retorna (peso_correto, peso_cobrado, tem_erro_peso).
        
        Lógica:
        - PESO_CORRETO = max(real, cubado) - O que DEVERIA ser usado
        - PESO_COBRADO = taxado se existir, senão usa peso_correto
        - TEM_ERRO = True se cobrado != correto
        """
        real = safe_float(peso_real)
        cubado = safe_float(peso_cubado)
        taxado = safe_float(peso_taxado)
        
        # Peso correto: sempre o maior entre real e cubado
        peso_correto = max(real, cubado) if (real > 0 or cubado > 0) else 1.0
        
        # Peso cobrado: usa taxado se disponível, senão assume correto
        peso_cobrado = taxado if taxado > 0 else peso_correto
        
        # Verifica erro
        tem_erro_peso = abs(peso_cobrado - peso_correto) > 0.5
        
        return (
            int(np.ceil(peso_correto)),
            int(np.ceil(peso_cobrado)),
            tem_erro_peso
        )

# ================================================================
# AUDITOR DE FRETE
# ================================================================

@dataclass
class ContextoLPU:
    df: pd.DataFrame
    kg_adicional: Dict[str, float]
    col_redespacho: str

class AuditorFrete:
    """Audita valores de frete baseado na tabela LPU."""
    
    def __init__(self, ctx: ContextoLPU, colunas_detectadas: dict):
        self.ctx = ctx
        self.colunas = list(ctx.df.columns)
        self.colunas_detectadas = colunas_detectadas
    
    def auditar_linha(self, row: pd.Series) -> pd.Series:
        """Audita uma linha do relatório."""
        # 1. EXTRAI PESOS
        peso_real = row.get(self.colunas_detectadas.get('peso_real'), 0)
        peso_cubado = row.get(self.colunas_detectadas.get('peso_cubado'), 0)
        peso_taxado = row.get(self.colunas_detectadas.get('peso_taxado'), 0)
        
        peso_correto, peso_cobrado, erro_peso = CalculadoraPeso.processar(
            peso_real, peso_cubado, peso_taxado
        )
        
        # 2. EXTRAI LOCALIZAÇÕES (com tratamento de UF dentro da cidade)
        localizacao = ExtratorLocalizacao.processar_linha(row, self.colunas_detectadas)
        
        origem_cidade = localizacao['origem_cidade']
        origem_uf = localizacao['origem_uf']
        destino_cidade = localizacao['destino_cidade']
        destino_uf = localizacao['destino_uf']
        
        # 3. CALCULA VALOR ESPERADO
        valor_lpu = self._calcular_valor_rota(
            origem_cidade, origem_uf,
            destino_cidade, destino_uf,
            peso_correto
        )
        
        # 4. COMPARA COM VALOR COBRADO
        frete_col = self.colunas_detectadas.get('frete_total')
        valor_cobrado = safe_float(row.get(frete_col, 0))
        diferenca = valor_cobrado - valor_lpu
        
        # 5. DETERMINA STATUS
        status, sugestao = self._analisar_divergencia(
            diferenca, valor_lpu, erro_peso, peso_correto, peso_cobrado
        )
        
        return pd.Series([
            peso_correto,
            peso_cobrado,
            round(valor_lpu, 2),
            round(diferenca, 2),
            status,
            sugestao
        ])
    
    def _calcular_valor_rota(self, orig_cid, orig_uf, dest_cid, dest_uf, peso):
        """Calcula valor do frete baseado na rota."""
        custo = 0.0
        
        # Origem
        if orig_cid and orig_cid not in HUB_CENTRAL:
            col, eh_interior = self._encontrar_coluna_destino(orig_cid, orig_uf)
            if col:
                custo += self._calcular_valor(peso, col)
                # Se é interior, soma a taxa de interior/redespacho
                if eh_interior:
                    custo += self._calcular_valor(peso, self.ctx.col_redespacho)
        
        # Destino
        if dest_cid and dest_cid not in HUB_CENTRAL:
            col, eh_interior = self._encontrar_coluna_destino(dest_cid, dest_uf)
            if col:
                custo += self._calcular_valor(peso, col)
                # Se é interior, soma a taxa de interior/redespacho
                if eh_interior:
                    custo += self._calcular_valor(peso, self.ctx.col_redespacho)
        
        # Se ambos são hub (SP local)
        if custo == 0:
            col_sp = next((c for c in self.colunas if "SP" in c and "CAPITAL" in c), self.colunas[0])
            custo = self._calcular_valor(peso, col_sp)
        
        return custo
    
    def _encontrar_coluna_destino(self, cidade, uf):
        """Encontra coluna da tabela LPU para determinada cidade.
        
        Retorna: (coluna, eh_interior)
        - coluna: nome da coluna encontrada
        - eh_interior: True se a cidade é do interior (não é capital/polo)
        """
        cidade_limpa = limpar_texto(cidade)
        
        # Busca direta - se encontra exatamente, é polo/capital
        for col in self.colunas:
            if cidade_limpa in limpar_texto(col):
                return (col, False)
        
        # Busca por capital - se encontra a capital, a cidade é interior
        capital = MAPA_UF_CAPITAL.get(uf)
        if capital:
            for col in self.colunas:
                if capital in limpar_texto(col):
                    return (col, True)  # Retorna que é interior
        
        # Fallback: redespacho (não soma interior novamente, já que é redespacho)
        return (self.ctx.col_redespacho, False)
    
    def _calcular_valor(self, peso, coluna):
        """Calcula valor baseado em peso e coluna da LPU."""
        if not coluna:
            return 0.0
        
        p_int = max(1, int(np.ceil(peso)))
        p_tab = min(p_int, 30)
        
        base = 0.0
        if p_tab in self.ctx.df.index:
            base = safe_float(self.ctx.df.loc[p_tab, coluna])
        
        if p_int > 30:
            adicional = self.ctx.kg_adicional.get(coluna, 0.0)
            return base + (p_int - 30) * adicional
        
        return base
    
    def _analisar_divergencia(self, diff, valor_lpu, erro_peso, peso_certo, peso_cobrado):
        """Analisa divergência e retorna status e sugestão."""
        if valor_lpu == 0:
            return "ERRO_CALCULO", "Não foi possível calcular valor da rota"
        
        if abs(diff) <= MARGEM_TOLERANCIA:
            if erro_peso:
                return "PESO_INCORRETO", f"Cobrado {peso_cobrado}kg ao invés de {peso_certo}kg"
            return "OK", "-"
        
        percentual = abs(diff / valor_lpu)
        
        if percentual > LIMITE_PERCENTUAL_CRITICO:
            tipo = "mais" if diff > 0 else "menos"
            msg = f"Pago a {tipo}: {percentual:.0%}"
            if erro_peso:
                msg += f" | Peso errado: {peso_cobrado}kg vs {peso_certo}kg"
            return "DIVERGENCIA_CRITICA", msg
        
        if erro_peso:
            return "PESO_INCORRETO", f"Peso incorreto: {peso_cobrado}kg vs {peso_certo}kg"
        
        return "OK", "-"

# ================================================================
# PROCESSADOR PRINCIPAL
# ================================================================

class ProcessadorAuditoria:
    """Coordena todo o processo de auditoria."""
    
    @staticmethod
    def processar(caminho_lpu: str, caminho_relatorio: str):
        """Executa auditoria completa."""
        # 1. CARREGA LPU
        ctx_lpu = ProcessadorAuditoria._carregar_lpu(caminho_lpu)
        
        # 2. CARREGA RELATÓRIO E DETECTA ESTRUTURA
        df_rel, colunas_detectadas = ProcessadorAuditoria._carregar_relatorio(caminho_relatorio)
        
        # 3. CRIA AUDITOR COM MAPEAMENTO
        auditor = AuditorFrete(ctx_lpu, colunas_detectadas)
        
        # 4. AUDITA
        resultado = df_rel.apply(auditor.auditar_linha, axis=1)
        resultado.columns = ['PESO_CORRETO', 'PESO_COBRADO', 'VALOR_LPU', 
                            'DIFERENCA', 'STATUS', 'SUGESTAO']
        
        # 5. MONTA RELATÓRIO FINAL
        df_final = pd.concat([df_rel, resultado], axis=1)
        
        return ProcessadorAuditoria._gerar_relatorio(df_final, colunas_detectadas)
    
    @staticmethod
    def _carregar_lpu(caminho: str) -> ContextoLPU:
        """Carrega e processa tabela LPU."""
        df = LeitorArquivo.carregar(caminho)
        
        # Encontra linha com "PESO"
        idx_peso = 0
        for i in range(min(20, len(df))):
            if "PESO" in str(df.iloc[i, 0]).upper():
                idx_peso = i
                break
        
        # Remove linhas antes do header
        df = df.iloc[idx_peso:].reset_index(drop=True)
        
        # Define header e deduplica
        df.columns = df.iloc[0]
        df = df.iloc[1:]
        df = LeitorArquivo.deduplica_colunas(df)
        
        # Define índice como peso
        col_peso = next((c for c in df.columns if 'PESO' in c and 'DUP' not in c), df.columns[0])
        df = df.set_index(col_peso)
        df.index = pd.to_numeric(df.index, errors='coerce')
        df = df.dropna(how='all')
        
        # Extrai kg adicional (última linha)
        kg_adicional = {c: safe_float(df.iloc[-1][c]) for c in df.columns}
        
        # Encontra coluna redespacho
        col_red = next((c for c in df.columns[::-1] if "REDESPACHO" in c or "INTERIOR" in c), 
                        df.columns[-1])
        
        return ContextoLPU(df, kg_adicional, col_red)
    
    @staticmethod
    def _carregar_relatorio(caminho: str):
        """Carrega e normaliza relatório de fretes."""
        df = LeitorArquivo.carregar(caminho)
        
        # Encontra cabeçalho
        palavras = ['PESO', 'CIDADE', 'FRETE', 'ORIGEM', 'DESTINO', 'REMETENTE', 'DESTINATARIO']
        idx_header = LeitorArquivo.encontrar_cabecalho(df, palavras)
        
        # Remove linhas antes do header
        df = df.iloc[idx_header:].reset_index(drop=True)
        df.columns = df.iloc[0]
        df = df.iloc[1:]
        
        # Deduplica colunas
        df = LeitorArquivo.deduplica_colunas(df)
        
        # DETECTA ESTRUTURA (novo sistema inteligente)
        colunas_detectadas = DetectorEstrutura.detectar(df.columns.tolist())
        
        return df, colunas_detectadas
    
    @staticmethod
    def _gerar_relatorio(df: pd.DataFrame, colunas_detectadas: dict):
        """Gera relatório formatado com totais."""
        # Monta lista de colunas para exportar (usa os nomes originais da planilha)
        cols_exportar = []
        
        # Adiciona colunas originais da planilha que foram detectadas
        if 'peso_real' in colunas_detectadas:
            cols_exportar.append(colunas_detectadas['peso_real'])
        if 'peso_cubado' in colunas_detectadas:
            cols_exportar.append(colunas_detectadas['peso_cubado'])
        if 'peso_taxado' in colunas_detectadas:
            cols_exportar.append(colunas_detectadas['peso_taxado'])
        if 'origem_cidade' in colunas_detectadas:
            cols_exportar.append(colunas_detectadas['origem_cidade'])
        if 'origem_uf' in colunas_detectadas:
            cols_exportar.append(colunas_detectadas['origem_uf'])
        if 'destino_cidade' in colunas_detectadas:
            cols_exportar.append(colunas_detectadas['destino_cidade'])
        if 'destino_uf' in colunas_detectadas:
            cols_exportar.append(colunas_detectadas['destino_uf'])
        if 'frete_total' in colunas_detectadas:
            cols_exportar.append(colunas_detectadas['frete_total'])
        
        # Adiciona colunas calculadas
        cols_exportar.extend(['PESO_CORRETO', 'PESO_COBRADO', 'VALOR_LPU', 
                            'DIFERENCA', 'STATUS', 'SUGESTAO'])
        
        # Filtra apenas colunas que existem no DataFrame
        cols_existentes = [c for c in cols_exportar if c in df.columns]
        df_export = df[cols_existentes].copy()
        
        # Calcula totais
        frete_col = colunas_detectadas.get('frete_total')
        if frete_col and frete_col in df_export.columns:
            total_pago = df_export[frete_col].apply(safe_float).sum()
        else:
            total_pago = 0.0
        
        total_devido = df_export['VALOR_LPU'].sum()
        total_diff = df_export['DIFERENCA'].sum()
        
        # Adiciona linha de total
        row_total = {col: np.nan for col in df_export.columns}
        
        # Preenche primeira coluna visível com "TOTAL GERAL"
        primeira_col = df_export.columns[0]
        row_total[primeira_col] = 'TOTAL GERAL'
        
        if frete_col and frete_col in df_export.columns:
            row_total[frete_col] = total_pago
        row_total['VALOR_LPU'] = total_devido
        row_total['DIFERENCA'] = total_diff
        
        df_export = pd.concat([df_export, pd.DataFrame([row_total])], ignore_index=True)
        
        # Aplica formatação
        def highlight(row):
            primeira_col_val = str(row.iloc[0]) if len(row) > 0 else ""
            if primeira_col_val == 'TOTAL GERAL':
                return ['background-color: #D3D3D3; font-weight: bold'] * len(row)
            if str(row.get('STATUS')) == 'DIVERGENCIA_CRITICA':
                return ['background-color: #FF5733; color: white'] * len(row)
            return [''] * len(row)
        
        styled = df_export.style.apply(highlight, axis=1)
        
        # Formata colunas monetárias
        format_dict = {
            'VALOR_LPU': formatar_moeda,
            'DIFERENCA': formatar_moeda
        }
        if frete_col and frete_col in df_export.columns:
            format_dict[frete_col] = formatar_moeda
        
        styled = styled.format(format_dict, na_rep="-")
        
        return styled, total_pago, total_devido, total_diff

# ================================================================
# INTERFACE GRÁFICA
# ================================================================

class AuditoriaFreteGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Auditoria de Frete v3.0")
        self.root.geometry("750x650")
        
        self.lpu_path = tk.StringVar()
        self.rel_path = tk.StringVar()
        self.resultado = None
        
        self._criar_interface()
    
    def _criar_interface(self):
        main = ttk.Frame(self.root, padding="20")
        main.pack(fill=tk.BOTH, expand=True)
        
        # Título
        tk.Label(main, text="🚚 AUDITORIA DE FRETE", 
                font=("Arial", 22, "bold"), fg="#1e3a8a").pack(pady=(0, 15))
        
        # Seleção LPU
        lpu_frame = ttk.LabelFrame(main, text="1. Tabela LPU", padding="10")
        lpu_frame.pack(fill=tk.X, pady=5)
        ttk.Entry(lpu_frame, textvariable=self.lpu_path, state="readonly", width=65).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(lpu_frame, text="📁 Selecionar", command=self._selecionar_lpu).pack(side=tk.LEFT)
        
        # Seleção Relatório
        rel_frame = ttk.LabelFrame(main, text="2. Relatório de Fretes", padding="10")
        rel_frame.pack(fill=tk.X, pady=5)
        ttk.Entry(rel_frame, textvariable=self.rel_path, state="readonly", width=65).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(rel_frame, text="📁 Selecionar", command=self._selecionar_rel).pack(side=tk.LEFT)
        
        # Botão processar
        self.btn_processar = tk.Button(main, text="⚙️ AUDITAR", font=("Arial", 13, "bold"), bg="#2563eb", fg="white", pady=12, command=self._processar)
        self.btn_processar.pack(fill=tk.X, pady=15)
        
        # Resultados
        result_frame = ttk.LabelFrame(main, text="📊 Resultados", padding="15")
        result_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.label_pago = tk.Label(result_frame, text="💰 Total Pago: Aguardando...", font=("Arial", 12, "bold"), anchor="w", bg="#dbeafe", fg="#1e40af", padx=12, pady=10)
        self.label_pago.pack(fill=tk.X, pady=5)
        
        self.label_lpu = tk.Label(result_frame, text="📋 Valor LPU: Aguardando...", font=("Arial", 12, "bold"), anchor="w", bg="#dbeafe", fg="#1e40af", padx=12, pady=10)
        self.label_lpu.pack(fill=tk.X, pady=5)
        
        self.label_diff = tk.Label(result_frame, text="📊 Diferença: Aguardando...", font=("Arial", 13, "bold"), anchor="w", bg="#fef3c7", fg="#92400e", padx=12, pady=12)
        self.label_diff.pack(fill=tk.X, pady=5)
        
        # Botão download
        self.btn_download = tk.Button(result_frame, text="💾 BAIXAR RELATÓRIO (EXCEL)", font=("Arial", 14, "bold"), bg="#16a34a", fg="white", pady=15, state=tk.DISABLED, command=self._baixar)
        self.btn_download.pack(fill=tk.X, pady=(15, 0))
        
        # Status
        self.status = tk.Label(main, text="Aguardando arquivos...", font=("Arial", 10), fg="#6b7280")
        self.status.pack(pady=(10, 0))
    
    def _selecionar_lpu(self):
        f = filedialog.askopenfilename(filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")])
        if f:
            self.lpu_path.set(f)
            self.status.config(text=f"✓ LPU: {os.path.basename(f)}", fg="green")
    
    def _selecionar_rel(self):
        f = filedialog.askopenfilename(filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")])
        if f:
            self.rel_path.set(f)
            self.status.config(text=f"✓ Relatório: {os.path.basename(f)}", fg="green")
    
    def _processar(self):
        if not self.lpu_path.get() or not self.rel_path.get():
            messagebox.showwarning("Atenção", "Selecione os dois arquivos!")
            return
        
        self.btn_processar.config(state=tk.DISABLED, text="⏳ Processando...")
        self.status.config(text="⏳ Processando...", fg="#ea580c")
        threading.Thread(target=self._processar_thread, daemon=True).start()
    
    def _processar_thread(self):
        try:
            resultado = ProcessadorAuditoria.processar(
                self.lpu_path.get(),
                self.rel_path.get()
            )
            self.resultado = resultado
            self.root.after(0, self._atualizar_resultados)
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Erro", str(e)))
            self.root.after(0, self._reset_botao)
    
    def _atualizar_resultados(self):
        _, pago, devido, diff = self.resultado
        
        self.label_pago.config(text=f"💰 Total Pago: {formatar_moeda(pago)}")
        self.label_lpu.config(text=f"📋 Valor LPU: {formatar_moeda(devido)}")
        
        cor = "#fecaca" if diff > 0 else "#bbf7d0"
        tipo = "PREJUÍZO" if diff > 0 else "ECONOMIA"
        sinal = "+" if diff > 0 else ""
        
        self.label_diff.config(
            text=f"📊 Diferença ({tipo}): {sinal}{formatar_moeda(diff)}",
            bg=cor
        )
        
        self.btn_download.config(state=tk.NORMAL)
        self.status.config(text="✅ Concluído!", fg="green")
        self._reset_botao()
    
    def _reset_botao(self):
        self.btn_processar.config(state=tk.NORMAL, text="⚙️ AUDITAR")
    
    def _baixar(self):
        f = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                        filetypes=[("Excel", "*.xlsx")])
        if f:
            self.resultado[0].to_excel(f, index=False, engine='openpyxl')
            messagebox.showinfo("Sucesso", "Relatório salvo com sucesso!")
            try:
                os.startfile(os.path.dirname(f))
            except:
                pass

# ================================================================
# MAIN
# ================================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = AuditoriaFreteGUI(root)
    root.mainloop()