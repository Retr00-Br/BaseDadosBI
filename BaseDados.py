import streamlit as st
import pandas as pd
import numpy as np
from supabase import create_client, Client
import datetime

st.set_page_config(
    page_title="Repositório de Expedição Multi-Abas",
    page_icon="📦",
    layout="wide"
)

st.title("📦 Ingestão Completa de Dados de Expedição")
st.caption("Sincronização automática de TODAS as abas da planilha com o Supabase (PostgreSQL).")


# -----------------------------------------------------------------------------
# Conexão com Supabase
# -----------------------------------------------------------------------------
@st.cache_resource
def init_supabase() -> Client:
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])


try:
    supabase = init_supabase()
except Exception as e:
    st.error("Erro ao conectar ao Supabase. Verifique suas credenciais em .streamlit/secrets.toml")
    st.stop()


# -----------------------------------------------------------------------------
# Funções de Higienização de Dados (Antierro de JSON / NaN)
# -----------------------------------------------------------------------------
def sanitize_value(val):
    """Converte valores NaN/NaT do Pandas/Numpy para None (null no JSON)."""
    if pd.isna(val) or val is np.nan:
        return None
    return val


def clean_records(records):
    """Limpa uma lista de dicionários substituindo qualquer NaN por None."""
    cleaned = []
    for row in records:
        cleaned_row = {k: sanitize_value(v) for k, v in row.items()}
        cleaned.append(cleaned_row)
    return cleaned


# -----------------------------------------------------------------------------
# Funções de Tratamento para Cada Aba
# -----------------------------------------------------------------------------

def process_sheet_base(xl_file):
    df = pd.read_excel(xl_file, sheet_name='BASE')
    df = df.rename(columns={
        'Nome Fantasia': 'nome_fantasia', 'Num.PV': 'num_pv', 'Num.NF': 'num_nf',
        'Vl.NF': 'vl_nf', 'Dt.Entrega': 'dt_entrega', 'TRANSPORTE': 'transporte',
        'DATA DE EMBARQUE': 'data_embarque'
    })
    
    # Remove registros sem número de nota fiscal
    df = df[pd.to_numeric(df['num_nf'], errors='coerce').notna()].copy()
    df['num_nf'] = df['num_nf'].astype(int)
    
    # Formatação de datas
    for col in ['dt_entrega', 'data_embarque']:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%Y-%m-%d')

    records = df.to_dict(orient='records')
    return clean_records(records)


def process_sheet_romaneios(xl_file):
    df = pd.read_excel(xl_file, sheet_name='ROMANEIOS')
    
    # Preenche cabeçalhos de romaneio mesclados para baixo
    for col in ['Nr. Romaneio', 'Dt. Romaneio', 'Desc.Veiculo', 'Placa Veiculo', 'Condutor']:
        if col in df.columns:
            df[col] = df[col].ffill()

    # Mapeia colunas conhecidas
    df = df.rename(columns={
        'Nr. Romaneio': 'nr_romaneio', 'Dt. Romaneio': 'dt_romaneio', 'Desc.Veiculo': 'desc_veiculo',
        'Placa Veiculo': 'placa_veiculo', 'Condutor': 'condutor', 'Unnamed: 5': 'filial',
        'Unnamed: 6': 'emissao', 'Unnamed: 7': 'cliente_loja', 'Unnamed: 8': 'nome_fantasia',
        'Unnamed: 9': 'num_pv', 'Unnamed: 10': 'num_nf', 'Unnamed: 11': 'vl_nf', 'Unnamed: 12': 'dt_entrega'
    })

    # Descarta linhas de cabeçalho repetidas e garante que num_nf seja numérico
    df = df[pd.to_numeric(df['num_nf'], errors='coerce').notna()].copy()
    df['num_nf'] = df['num_nf'].astype(int)

    # Cria ID único composto
    df['id_romaneio_nf'] = df['nr_romaneio'].astype(str) + "_" + df['num_nf'].astype(str)

    # Formatação de datas
    df['dt_romaneio'] = pd.to_datetime(df['dt_romaneio'], errors='coerce').dt.strftime('%Y-%m-%d')
    df['dt_entrega'] = pd.to_datetime(df['dt_entrega'], errors='coerce').dt.strftime('%Y-%m-%d')
    df['emissao'] = pd.to_datetime(df['emissao'], errors='coerce').dt.strftime('%Y-%m-%d %H:%M:%S')

    # 🚨 SOLUÇÃO DO ERRO: Mantém APENAS as colunas válidas no banco de dados
    valid_columns = [
        'id_romaneio_nf', 'nr_romaneio', 'dt_romaneio', 'desc_veiculo', 
        'placa_veiculo', 'condutor', 'filial', 'emissao', 'cliente_loja', 
        'nome_fantasia', 'num_pv', 'num_nf', 'vl_nf', 'dt_entrega'
    ]
    
    # Filtra apenas as colunas que existem na lista acima
    cols_to_keep = [col for col in valid_columns if col in df.columns]
    df = df[cols_to_keep]

    records = df.to_dict(orient='records')
    return clean_records(records)


def process_sheet_consolidado(xl_file):
    df = pd.read_excel(xl_file, sheet_name='CONSOLIDADO')
    df = df.rename(columns={
        'DATA DE ROMANEIO': 'data_romaneio', 'TOTAL DE NOTAS EXPEDIDAS': 'total_notas_expedidas',
        'VALOR TOTAL EMBARCADO': 'valor_total_embarcado', 'EMB. TERCEIROS': 'emb_terceiros',
        'EMB. CARRO PROPRIO': 'emb_carro_proprio', 'EMB. TRANSPORTADORAS': 'emb_transportadoras',
        'EMB. CORREIOS': 'emb_correios', 'RETIRA': 'retira', 'EMB. MOTOBOY': 'emb_motoboy',
        'EMB. LALAMOVE': 'emb_lalamove'
    })
    df['data_romaneio'] = pd.to_datetime(df['data_romaneio'], errors='coerce').dt.strftime('%Y-%m-%d')
    df = df[df['data_romaneio'].notna()]

    records = df.to_dict(orient='records')
    return clean_records(records)


def process_sheet_planilha6(xl_file):
    df = pd.read_excel(xl_file, sheet_name='Planilha6')
    df = df.rename(columns={
        'Rótulos de Linha': 'data_rotulo', 'Soma de TOTAL DE NOTAS EXPEDIDAS': 'total_notas_expedidas',
        'Soma de EMB. CARRO PROPRIO': 'emb_carro_proprio', 'Soma de EMB. TERCEIROS': 'emb_terceiros',
        'Soma de EMB. TRANSPORTADORAS': 'emb_transportadoras', 'Soma de EMB. CORREIOS': 'emb_correios',
        'Soma de RETIRA': 'retira', 'Soma de EMB. MOTOBOY': 'emb_motoboy',
        'Soma de EMB. LALAMOVE': 'emb_lalamove', 'Soma de VALOR TOTAL EMBARCADO': 'valor_total_embarcado'
    })
    df['data_rotulo'] = pd.to_datetime(df['data_rotulo'], errors='coerce').dt.strftime('%Y-%m-%d')
    df = df[df['data_rotulo'].notna()]

    records = df.to_dict(orient='records')
    return clean_records(records)


# -----------------------------------------------------------------------------
# Interface de Upload
# -----------------------------------------------------------------------------
uploaded_file = st.file_uploader("Upload da Planilha Completa de Expedição (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        rec_base = process_sheet_base(uploaded_file)
        rec_romaneios = process_sheet_romaneios(uploaded_file)
        rec_consolidado = process_sheet_consolidado(uploaded_file)
        rec_resumo = process_sheet_planilha6(uploaded_file)

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Aba BASE", f"{len(rec_base)} NFs")
        col2.metric("Aba ROMANEIOS", f"{len(rec_romaneios)} itens")
        col3.metric("Aba CONSOLIDADO", f"{len(rec_consolidado)} dias")
        col4.metric("Aba PLANILHA6", f"{len(rec_resumo)} resumos")

        if st.button("🚀 Sincronizar TODAS as Abas no Supabase", type="primary"):
            with st.spinner("Realizando UPSERT nas 4 tabelas do banco..."):
                # Upsert em cada tabela
                supabase.table("expedicao_base").upsert(rec_base, on_conflict="num_nf").execute()
                supabase.table("expedicao_romaneios").upsert(rec_romaneios, on_conflict="id_romaneio_nf").execute()
                supabase.table("expedicao_consolidado").upsert(rec_consolidado, on_conflict="data_romaneio").execute()
                supabase.table("expedicao_resumo").upsert(rec_resumo, on_conflict="data_rotulo").execute()

                st.success(
                    "✅ Sincronização multi-aba concluída com sucesso! Todas as tabelas foram atualizadas sem duplicatas.")

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {str(e)}")
