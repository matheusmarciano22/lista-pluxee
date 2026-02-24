import streamlit as st
import pandas as pd
import re
from datetime import datetime
import io
import os
from unidecode import unidecode
from dateutil import parser
from dateutil.relativedelta import relativedelta
from thefuzz import process
import openpyxl
import requests

# Tenta importar o módulo para leitura de ficheiros Word
try:
    import docx
except ImportError:
    pass

# ==========================================
# LÓGICA DE TRATAMENTO DE DADOS 
# ==========================================

def formatar_nome_pluxee(nome_bruto, limite=40):
    """Abrevia nomes do meio, mantendo Primeiro e Último intactos."""
    nome = unidecode(str(nome_bruto)).upper().strip()
    if len(nome) <= limite:
        return nome
    partes = nome.split()
    if len(partes) <= 2:
        return nome[:limite]
    primeiro = partes[0]
    ultimo = partes[-1]
    meio = partes[1:-1]
    for i in range(len(meio)):
        if len(meio[i]) > 2: 
            meio[i] = meio[i][0] + "."
        tentativa = " ".join([primeiro] + meio + [ultimo])
        if len(tentativa) <= limite:
            return tentativa
    extremo = f"{primeiro} {ultimo}"
    return extremo[:limite]

def formatar_local(texto_bruto, limite=30):
    """Abrevia bairros e cidades para caber no limite da Pluxee."""
    if pd.isna(texto_bruto): return ""
    texto = unidecode(str(texto_bruto)).upper().strip()
    if len(texto) <= limite:
        return texto
    partes = texto.split()
    if len(partes) <= 2:
        return texto[:limite]
    primeiro = partes[0]
    ultimo = partes[-1]
    meio = partes[1:-1]
    for i in range(len(meio)):
        if len(meio[i]) > 2: 
            meio[i] = meio[i][0] + "."
        tentativa = " ".join([primeiro] + meio + [ultimo])
        if len(tentativa) <= limite:
            return tentativa
    return texto[:limite]

def limpar_cpf(cpf_bruto):
    """Garante CPF com 11 dígitos numéricos."""
    if pd.isna(cpf_bruto): return ""
    cpf_limpo = re.sub(r'\D', '', str(cpf_bruto))
    return cpf_limpo.zfill(11)

def converter_data(data_bruta, data_padrao):
    """Converte qualquer formato de data para DD/MM/AAAA."""
    if pd.isna(data_bruta) or str(data_bruta).strip() == "":
        return data_padrao
    try:
        dt = parser.parse(str(data_bruta), dayfirst=True, fuzzy=True)
        return dt.strftime('%d/%m/%Y')
    except:
        return data_padrao

# ==========================================
# INTERFACE E LIGAÇÃO API
# ==========================================
st.set_page_config(page_title="Gerador Pluxee Oficial", page_icon="💳", layout="wide")
st.title("💳 Máquina de Geração - Pluxee PLANSIP3C")

col1, col2 = st.columns([1, 2])

# Inicializa o estado para evitar o erro de Series ambígua
if "config_rh" not in st.session_state:
    st.session_state.config_rh = None
if "razao_social" not in st.session_state:
    st.session_state.razao_social = "Cliente_Novo"

with col1:
    st.markdown("### ⚙️ Dados Fixos do Pedido")
    st.markdown("**🔐 Ligação ao Lovable**")
    email_lovable = st.text_input("E-mail Lovable", "seu@email.com")
    senha_lovable = st.text_input("Password Lovable", type="password")
    
    @st.cache_data(ttl=300)
    def carregar_clientes_lovable(email, senha):
        if email == "seu@email.com" or senha == "":
            return pd.DataFrame()
        url_base = "https://rlbcxlgqfvcloieywfiq.supabase.co"
        anon_key = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJsYmN4bGdxZnZjbG9pZXl3ZmlxIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Njk1MzA3NDIsImV4cCI6MjA4NTEwNjc0Mn0.Rj6JhvcGscQ6quQ9K6QwDdiHjbh2pHOZVS7gwDQHDV0"
        try:
            auth = requests.post(f"{url_base}/auth/v1/token?grant_type=password",
                json={"email": email, "password": senha},
                headers={"apikey": anon_key, "Content-Type": "application/json"})
            if auth.status_code != 200:
                return pd.DataFrame()
            token = auth.json()["access_token"]
            resp = requests.get(f"{url_base}/functions/v1/api-vendas",
                headers={"Authorization": f"Bearer {token}", "apikey": anon_key})
            if resp.status_code == 200:
                lista = resp.json().get("vendas", [])
                if lista:
                    df = pd.DataFrame(lista)
                    return df.rename(columns={"cliente_nome": "razao_social", "responsavel_pedido": "responsavel", "estado": "uf"})
        except Exception as e:
            st.error(f"Erro de ligação: {e}")
        return pd.DataFrame()

    df_clientes = carregar_clientes_lovable(email_lovable, senha_lovable)
    st.markdown("---")

    if not df_clientes.empty:
        st.success(f"✅ {len(df_clientes)} vendas encontradas.")
        empresa_sel = st.selectbox("🔍 Selecione o Cliente", df_clientes['razao_social'].tolist())
        
        # Converte a Series para um dicionário Python comum para evitar erro de ambiguidade
        dados_e = df_clientes[df_clientes['razao_social'] == empresa_sel].iloc[0].to_dict()
        st.session_state.razao_social = empresa_sel
        
        st.session_state.config_rh = {
            'Local de entrega': st.text_input("Local de Entrega", "MATRIZ"),
            'CEP': st.text_input("CEP", str(dados_e.get('endereco_cep', '00000-000'))),
            'Endereço': st.text_input("Endereço", str(dados_e.get('endereco', 'RUA EXEMPLO'))),
            'Número': st.text_input("Número", str(dados_e.get('numero', '100'))),
            'Complemento': st.text_input("Complemento", str(dados_e.get('endereco_complemento', ''))),
            'Referência': st.text_input("Referência", ""),
            'Bairro': st.text_input("Bairro", formatar_local(dados_e.get('endereco_bairro', 'CENTRO')), max_chars=30),
            'Cidade': st.text_input("Cidade", formatar_local(dados_e.get('cidade', 'SÃO PAULO')), max_chars=30),
            'UF': st.text_input("UF", str(dados_e.get('uf', 'SP')).upper(), max_chars=2),
            'Responsável': st.text_input("Responsável", str(dados_e.get('responsavel', ''))),
            'DDD': st.text_input("DDD", "11", max_chars=2),
            'Telefone': st.text_input("Telefone", "999999999"),
            'Email': st.text_input("Email", ""),
            'Porta_a_Porta': st.selectbox("Porta a Porta?", ["Não", "Sim"])
        }
    else:
        st.info("Faça login para carregar dados do Lovable.")

with col2:
    st.markdown("### 📥 Ficheiro de Funcionários")
    arq = st.file_uploader("Suba Excel (.xlsx, .xls), CSV ou Word (.docx)", type=["xlsx", "xls", "csv", "docx"])
    
    if arq:
        template_path = "PLANSIP3C_NOVA.xlsx"
        if not os.path.exists(template_path):
            st.error("Arquivo PLANSIP3C_NOVA.xlsx não encontrado no GitHub.")
            st.stop()

        try:
            # --- LEITOR WORD ---
            if arq.name.endswith('.docx'):
                doc = docx.Document(arq)
                linhas = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
                dados_ex = []
                p_atual = {'nome': '', 'cpf': '', 'datas': []}
                for l in linhas:
                    nums = re.sub(r'\D', '', l)
                    m_data = re.search(r'\d{2}/\d{2}/\d{2,4}', l)
                    if m_data: p_atual['datas'].append(m_data.group())
                    elif 9 <= len(nums) <= 14:
                        if not p_atual['cpf']: p_atual['cpf'] = nums
                    else:
                        if p_atual['nome'] and p_atual['cpf']:
                            d_nasc = ""
                            if p_atual['datas']:
                                validas = []
                                for d in p_atual['datas']:
                                    try: validas.append((parser.parse(d, dayfirst=True), d))
                                    except: pass
                                if validas:
                                    validas.sort(key=lambda x: x[0])
                                    d_nasc = validas[0][1]
                            dados_ex.append({'nome': p_atual['nome'], 'cpf': p_atual['cpf'], 'nascimento': d_nasc})
                            p_atual = {'nome': l, 'cpf': '', 'datas': []}
                        elif not p_atual['nome']: p_atual['nome'] = l
                
                if p_atual['nome'] and p_atual['cpf']:
                    dv = []
                    for d in p_atual['datas']:
                        try: dv.append((parser.parse(d, dayfirst=True), d))
                        except: pass
                    dn = dv[0][1] if dv else ""
                    dados_ex.append({'nome': p_atual['nome'], 'cpf': p_atual['cpf'], 'nascimento': dn})
                
                data_rows = pd.DataFrame(dados_ex)
                c_nome, c_cpf, c_nasc = 'nome', 'cpf', 'nascimento'

            # --- LEITOR EXCEL/CSV ---
            else:
                if arq.name.endswith('.csv'): df_cli = pd.read_csv(arq, header=None)
                else: df_cli = pd.read_excel(arq, header=None)
                
                start_row = 0
                for i, row in df_cli.head(20).iterrows():
                    l_txt = str(row.values).lower()
                    if 'nome' in l_txt and 'cpf' in l_txt:
                        start_row = i
                        break
                headers = [str(c).lower().strip() for c in df_cli.iloc[start_row]]
                data_rows = df_cli.iloc[start_row+1:].reset_index(drop=True)
                data_rows.columns = headers
                c_nome = headers[headers.index(process.extractOne("nome", headers)[0])]
                c_cpf = headers[headers.index(process.extractOne("cpf", headers)[0])]
                m_nasc = process.extractOne("nascimento", headers)
                c_nasc = headers[headers.index(m_nasc[0])] if m_nasc[1] >= 70 else None

            # --- PROCESSAMENTO ---
            # Verificação segura se config_rh existe no session_state
            if st.session_state.config_rh is not None:
                if st.button("🚀 Gerar Planilha Pluxee Oficial", use_container_width=True):
                    wb = openpyxl.load_workbook(template_path)
                    ws = wb["Dados dos Beneficiários"]
                    r_idx = 8
                    dt_cred = (datetime.now() + relativedelta(months=1)).strftime('%d/%m/%Y')
                    
                    config = st.session_state.config_rh
                    
                    for _, r in data_rows.iterrows():
                        v_n, v_c = r.get(c_nome), r.get(c_cpf)
                        if pd.isna(v_n) or str(v_n).strip() == "" or pd.isna(v_c): continue
                        
                        nf, cpfl = formatar_nome_pluxee(v_n), limpar_cpf(v_c)
                        nasc = converter_data(r.get(c_nasc), "01/01/1980") if c_nasc else "01/01/1980"

                        for p_code in ["6001 - Carteira Refeição", "6002 - Carteira Alimentação"]:
                            ws.cell(row=r_idx, column=2, value="Ativo")
                            ws.cell(row=r_idx, column=3, value=nf)
                            ws.cell(row=r_idx, column=6, value=nf)
                            ws.cell(row=r_idx, column=4, value=cpfl)
                            ws.cell(row=r_idx, column=5, value=nasc)
                            ws.cell(row=r_idx, column=11, value="001 - Pedido Normal")
                            ws.cell(row=r_idx, column=12, value=p_code)
                            ws.cell(row=r_idx, column=13, value=0)
                            ws.cell(row=r_idx, column=14, value=dt_cred)
                            ws.cell(row=r_idx, column=16, value=config.get('Local de entrega'))
                            ws.cell(row=r_idx, column=17, value=config.get('CEP'))
                            ws.cell(row=r_idx, column=18, value=config.get('Endereço'))
                            ws.cell(row=r_idx, column=19, value=config.get('Número'))
                            ws.cell(row=r_idx, column=20, value=config.get('Complemento'))
                            ws.cell(row=r_idx, column=22, value=config.get('Bairro'))
                            ws.cell(row=r_idx, column=23, value=config.get('Cidade'))
                            ws.cell(row=r_idx, column=24, value=config_rh.get('UF', 'SP'))
                            ws.cell(row=r_idx, column=25, value=config.get('Responsável'))
                            ws.cell(row=r_idx, column=26, value=config.get('DDD'))
                            ws.cell(row=r_idx, column=27, value=config.get('Telefone'))
                            ws.cell(row=r_idx, column=28, value=config.get('Email'))
                            ws.cell(row=r_idx, column=29, value=config.get('Porta_a_Porta'))
                            r_idx += 1
                    
                    buf = io.BytesIO()
                    wb.save(buf)
                    buf.seek(0)
                    st.success("✅ Processado com sucesso!")
                    st.download_button(label="⬇️ Baixar PLANSIP3C", data=buf, 
                        file_name=f"PLANSIP3C_{re.sub(r'[^A-Za-z0-9]', '', st.session_state.razao_social)}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            else:
                st.warning("⚠️ Selecione um cliente no menu à esquerda primeiro.")
        except Exception as e:
            st.error(f"Erro no processamento: {e}")
