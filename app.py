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

try:
    import docx
except ImportError:
    pass

# ==========================================
# LÓGICA DE TRATAMENTO DE DADOS 
# ==========================================

def formatar_nome_pluxee(nome_bruto, limite=40):
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
    extremo = f"{primeiro} {ultimo}"
    return extremo[:limite]

def limpar_cpf(cpf_bruto):
    if pd.isna(cpf_bruto): return ""
    cpf_limpo = re.sub(r'\D', '', str(cpf_bruto))
    return cpf_limpo.zfill(11)

def converter_data(data_bruta, data_padrao):
    if pd.isna(data_bruta) or str(data_bruta).strip() == "":
        return data_padrao
    try:
        dt = parser.parse(str(data_bruta), dayfirst=True, fuzzy=True)
        return dt.strftime('%d/%m/%Y')
    except:
        return data_padrao

# ==========================================
# INTERFACE E PROCESSAMENTO
# ==========================================
st.set_page_config(page_title="Gerador Pluxee Oficial", page_icon="💳", layout="wide")
st.title("💳 Máquina de Geração - Pluxee PLANSIP3C")

col1, col2 = st.columns([1, 2])

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
            auth = requests.post(
                f"{url_base}/auth/v1/token?grant_type=password",
                json={"email": email, "password": senha},
                headers={"apikey": anon_key, "Content-Type": "application/json"}
            )
            
            if auth.status_code != 200:
                st.error("❌ E-mail ou password do Lovable incorretos.")
                return pd.DataFrame()
                
            token = auth.json()["access_token"]
            
            resp = requests.get(
                f"{url_base}/functions/v1/api-vendas",
                headers={
                    "Authorization": f"Bearer {token}",
                    "apikey": anon_key
                }
            )
            
            if resp.status_code == 200:
                vendas_json = resp.json()
                lista_vendas = vendas_json.get("vendas", [])
                
                if len(lista_vendas) > 0:
                    df = pd.DataFrame(lista_vendas)
                    df = df.rename(columns={
                        "cliente_nome": "razao_social",
                        "responsavel_pedido": "responsavel",
                        "estado": "uf"
                    })
                    return df
                else:
                    st.warning("Nenhuma venda encontrada na API.")
            else:
                st.error(f"Erro ao procurar vendas na API: {resp.status_code}")
                
        except Exception as e:
            st.error(f"Falha na ligação: {e}")
            
        return pd.DataFrame()

    df_clientes = carregar_clientes_lovable(email_lovable, senha_lovable)

    st.markdown("---")

    if not df_clientes.empty and 'razao_social' in df_clientes.columns:
        st.success(f"✅ Ligado! {len(df_clientes)} vendas encontradas.")
        lista_empresas = df_clientes['razao_social'].tolist()
        empresa_selecionada = st.selectbox("🔍 Selecione o Cliente (Venda Lovable)", lista_empresas)

        dados_empresa = df_clientes[df_clientes['razao_social'] == empresa_selecionada].iloc[0]

        st.markdown("**Endereço e Contato (RH) - Preenchimento Automático**")
        
        bairro_formatado = formatar_local(dados_empresa.get('endereco_bairro', 'CENTRO'), 30)
        cidade_formatada = formatar_local(dados_empresa.get('cidade', 'SÃO PAULO'), 30)
        
        config_rh = {
            'Local de entrega': st.text_input("Local de Entrega", "MATRIZ"),
            'CEP': st.text_input("CEP", str(dados_empresa.get('endereco_cep', '00000-000'))),
            'Endereço': st.text_input("Endereço", str(dados_empresa.get('endereco', 'RUA EXEMPLO'))),
            'Número': st.text_input("Número", str(dados_empresa.get('numero', '100'))),
            'Complemento': st.text_input("Complemento", str(dados_empresa.get('endereco_complemento', ''))),
            'Referência': st.text_input("Referência", ""),
            'Bairro': st.text_input("Bairro", bairro_formatado, max_chars=30),
            'Cidade': st.text_input("Cidade", cidade_formatada, max_chars=30),
            'UF': st.text_input("UF", str(dados_empresa.get('uf', 'SP')).upper(), max_chars=2),
            'Responsável': st.text_input("Nome do Responsável", str(dados_empresa.get('responsavel', ''))),
            'DDD': st.text_input("DDD", "11", max_chars=2),
            'Telefone': st.text_input("Telefone", "999999999"),
            'Email': st.text_input("Email", ""),
            'Porta_a_Porta': st.selectbox("Entrega Porta a Porta?", ["Não", "Sim"])
        }
        razao_social = empresa_selecionada
    else:
        st.info("👆 Faça o login acima com a sua conta do Lovable.")
        razao_social = "Cliente_Novo"
        config_rh = {
            'Local de entrega': "MATRIZ", 'CEP': "00000-000", 'Endereço': "RUA EXEMPLO",
            'Número': "100", 'Complemento': "", 'Referência': "", 'Bairro': "CENTRO",
            'Cidade': "SÃO PAULO", 'UF': "SP", 'Responsável': "", 'DDD': "11",
            'Telefone': "999999999", 'Email': "", 'Porta_a_Porta': "Não"
        }

with col2:
    st.markdown("### 📥 Ficheiro de Funcionários")
    arquivo_cliente = st.file_uploader("Suba a folha de cálculo (.xlsx, .xls, .csv) ou Word (.docx)", type=["xlsx", "xls", "csv", "docx"])
    
    if arquivo_cliente:
        template_path = "PLANSIP3C_NOVA.xlsx"
        if not os.path.exists(template_path):
            st.error(f"⚠️ O ficheiro original '{template_path}' não está no servidor!")
            st.stop()

        try:
            # Lógica de Leitura DOCX (Word)
            if arquivo_cliente.name.endswith('.docx'):
                doc = docx.Document(arquivo_cliente)
                linhas_texto = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
                dados_extraidos = []
                pessoa_atual = {'nome': '', 'cpf': '', 'datas_encontradas': []}
                
                for linha in linhas_texto:
                    numeros = re.sub(r'\D', '', linha)
                    match_data = re.search(r'\d{2}/\d{2}/\d{2,4}', linha)
                    if match_data:
                        pessoa_atual['datas_encontradas'].append(match_data.group())
                    elif 9 <= len(numeros) <= 14:
                        if not pessoa_atual['cpf']: pessoa_atual['cpf'] = numeros
                    else:
                        if pessoa_atual['nome'] and pessoa_atual['cpf']:
                            data_nascimento = ""
                            if pessoa_atual['datas_encontradas']:
                                dv = []
                                for d in pessoa_atual['datas_encontradas']:
                                    try: dv.append((parser.parse(d, dayfirst=True), d))
                                    except: pass
                                if dv:
                                    dv.sort(key=lambda x: x[0])
                                    data_nascimento = dv[0][1]
                            dados_extraidos.append({'nome': pessoa_atual['nome'], 'cpf': pessoa_atual['cpf'], 'nascimento': data_nascimento})
                            pessoa_atual = {'nome': linha, 'cpf': '', 'datas_encontradas': []}
                        elif not pessoa_atual['nome']: pessoa_atual['nome'] = linha

                if pessoa_atual['nome'] and pessoa_atual['cpf']:
                    dv = []
                    for d in pessoa_atual['datas_encontradas']:
                        try: dv.append((parser.parse(d, dayfirst=True), d))
                        except: pass
                    dn = dv[0][1] if dv else ""
                    dados_extraidos.append({'nome': pessoa_atual['nome'], 'cpf': pessoa_atual['cpf'], 'nascimento': dn})
                
                data_rows = pd.DataFrame(dados_extraidos)
                c_nome, c_cpf, c_nasc = 'nome', 'cpf', 'nascimento'
            
            # Lógica de Leitura Excel/CSV
            else:
                if arquivo_cliente.name.endswith('.csv'):
                    df_cli = pd.read_csv(arquivo_cliente, header=None)
                else:
                    # Suporte a .xls (via xlrd) e .xlsx
                    df_cli = pd.read_excel(arquivo_cliente, header=None)
                
                start_row = 0
                for i, row in df_cli.head(20).iterrows():
                    linha_txt = str(row.values).lower()
                    if 'nome' in linha_txt and 'cpf' in linha_txt:
                        start_row = i
                        break
                
                headers = [str(c).lower().strip() for c in df_cli.iloc[start_row]]
                data_rows = df_cli.iloc[start_row+1:].reset_index(drop=True)
                data_rows.columns = headers
                c_nome = headers[headers.index(process.extractOne("nome", headers)[0])]
                c_cpf = headers[headers.index(process.extractOne("cpf", headers)[0])]
                match_nasc = process.extractOne("nascimento", headers)
                c_nasc = headers[headers.index(match_nasc[0])] if match_nasc[1] >= 70 else None

            # Processamento Final
            if st.button("🚀 Processar e Gerar Pedido Oficial", use_container_width=True):
                wb = openpyxl.load_workbook(template_path)
                ws = wb["Dados dos Beneficiários"]
                row_idx = 8
                data_credito = (datetime.now() + relativedelta(months=1)).strftime('%d/%m/%Y')

                for _, r in data_rows.iterrows():
                    v_nome, v_cpf = r.get(c_nome), r.get(c_cpf)
                    if pd.isna(v_nome) or str(v_nome).strip() == "" or pd.isna(v_cpf): continue
                    
                    nf = formatar_nome_pluxee(v_nome)
                    cpfl = limpar_cpf(v_cpf)
                    nasc = converter_data(r.get(c_nasc), "01/01/1980") if c_nasc else "01/01/1980"

                    for p_code in ["6001 - Carteira Refeição", "6002 - Carteira Alimentação"]:
                        ws.cell(row=row_idx, column=2, value="Ativo")
                        ws.cell(row=row_idx, column=3, value=nf)
                        ws.cell(row=row_idx, column=6, value=nf)
                        ws.cell(row=row_idx, column=4, value=cpfl)
                        ws.cell(row=row_idx, column=5, value=nasc)
                        ws.cell(row=row_idx, column=11, value="001 - Pedido Normal")
                        ws.cell(row=row_idx, column=12, value=p_code)
                        ws.cell(row=row_idx, column=13, value=0)
                        ws.cell(row=row_idx, column=14, value=data_credito)
                        ws.cell(row=row_idx, column=16, value=config_rh['Local de entrega'])
                        ws.cell(row=row_idx, column=17, value=config_rh['CEP'])
                        ws.cell(row=row_idx, column=18, value=config_rh['Endereço'])
                        ws.cell(row=row_idx, column=19, value=config_rh['Número'])
                        ws.cell(row=row_idx, column=22, value=config_rh['Bairro'])
                        ws.cell(row=row_idx, column=23, value=config_rh['Cidade'])
                        ws.cell(row=row_idx, column=24, value=config_rh['UF'])
                        ws.cell(row=row_idx, column=25, value=config_rh['Responsável'])
                        ws.cell(row=row_idx, column=28, value=config_rh['Email'])
                        ws.cell(row=row_idx, column=29, value=config_rh['Porta_a_Porta'])
                        row_idx += 1

                buf = io.BytesIO()
                wb.save(buf)
                buf.seek(0)
                
                st.success(f"✅ Pedido finalizado!")
                st.download_button(
                    label="⬇️ Download da Planilha PLANSIP3C",
                    data=buf,
                    file_name=f"PLANSIP3C_{re.sub(r'[^A-Za-z0-9]', '', razao_social)}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"Erro: {e}")
