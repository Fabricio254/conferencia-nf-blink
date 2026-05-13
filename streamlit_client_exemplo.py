"""
Exemplo de Cliente Streamlit que consome a API FastAPI
Para usar: colocar em Streamlit Cloud e configurar secrets.toml
"""

import streamlit as st
import requests
import pandas as pd
from datetime import datetime, timedelta

# ═════════════════════════════════════════════════════════════════════════════
# CONFIGURAÇÃO
# ═════════════════════════════════════════════════════════════════════════════

# Ler configurações do Streamlit secrets
try:
    if "api" in st.secrets:
        API_URL = st.secrets["api"]["url"]
        API_TOKEN = st.secrets["api"]["token"]
    else:
        API_URL = st.secrets.get("api_url", None)
        API_TOKEN = st.secrets.get("api_token", None)
except Exception:
    API_URL = None
    API_TOKEN = None

# Headers de autenticação
HEADERS = {"Authorization": f"Bearer {API_TOKEN}"} if API_TOKEN else {}

# ═════════════════════════════════════════════════════════════════════════════
# FUNÇÕES DE CLIENTE API
# ═════════════════════════════════════════════════════════════════════════════

@st.cache_resource
def health_check():
    """Verifica se a API está disponível."""
    try:
        response = requests.get(f"{API_URL}/api/health", timeout=5)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        st.error(f"❌ Erro ao conectar API: {e}")
        return None

def buscar_nfs_entrada(data_ini: str, data_fim: str, loja_id: int = None):
    """Busca NFs de entrada."""
    try:
        params = {
            "data_ini": data_ini,
            "data_fim": data_fim,
        }
        if loja_id:
            params["loja_id"] = loja_id
        
        response = requests.get(
            f"{API_URL}/api/nfs-entrada",
            params=params,
            headers=HEADERS,
            timeout=30,
        )
        response.raise_for_status()
        return response.json()
    except requests.exceptions.HTTPError as e:
        if e.response.status_code == 401:
            st.error("❌ Token de autenticação inválido")
        elif e.response.status_code == 403:
            st.error("❌ Acesso negado")
        else:
            st.error(f"❌ Erro HTTP: {e}")
        return []
    except requests.exceptions.RequestException as e:
        st.error(f"❌ Erro na requisição: {e}")
        return []

def buscar_nfs_saida(data_ini: str, data_fim: str, loja_id: int = None):
    """Busca NFs de saída."""
    try:
        params = {
            "data_ini": data_ini,
            "data_fim": data_fim,
        }
        if loja_id:
            params["loja_id"] = loja_id
        
        response = requests.get(
            f"{API_URL}/api/nfs-saida",
            params=params,
            headers=HEADERS,
            timeout=30,
        )
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        st.error(f"❌ Erro ao buscar NFs de saída: {e}")
        return []

def buscar_lojas():
    """Lista todas as lojas."""
    try:
        response = requests.get(
            f"{API_URL}/api/lojas",
            headers=HEADERS,
            timeout=10,
        )
        response.raise_for_status()
        data = response.json()
        return data.get("lojas", [])
    except requests.exceptions.RequestException as e:
        st.error(f"❌ Erro ao buscar lojas: {e}")
        return []

# ═════════════════════════════════════════════════════════════════════════════
# INTERFACE STREAMLIT
# ═════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="Conferência NF — Via API",
    page_icon="📋",
    layout="wide",
)

st.title("📋 Conferência de Notas Fiscais — Via API")
st.markdown("**Acessando Firebird através da API FastAPI**")

# ─── VERIFICAR SECRETS ─────────────────────────────────────────────────────────
if not API_URL or not API_TOKEN:
    st.error("⚠️ **Secrets não configurados!**")
    st.warning("""
### Para configurar:

1. Acesse: **https://share.streamlit.io/**
2. Clique na app `conferencia-nf-blink`
3. Vá em Settings (⚙️) → **Secrets**
4. Cole o seguinte (formato TOML):

```toml
[api]
url = "https://conferencia-nf-blink.ngrok-free.dev"
token = "seu_token_super_secreto_aqui"
```

5. Clique "Save"
6. Recarregue a página (F5)

**ANTES de fazer deploy, você PRECISA:**
- Iniciar a API localmente: `Iniciar_API_com_ngrok.bat`
- Isso expõe a API com ngrok e gera a URL pública
    """)
    st.stop()

# Health Check
with st.spinner("Verificando conexão com API..."):
    health = health_check()

if health and health.get("status") == "OK":
    st.success(f"✅ {health['message']}")
else:
    st.error(f"❌ Não foi possível conectar à API em {API_URL}")
    st.info("""
**Verifique:**
- ✓ A API está rodando (`Iniciar_API_com_ngrok.bat`)
- ✓ ngrok está gerando a URL pública
- ✓ O token está correto nos secrets
- ✓ A URL da API está acessível (teste em: https://share.streamlit.io/fabricio254/conferencia-nf-blink)
    """)
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# SELEÇÃO DE FILTROS
# ─────────────────────────────────────────────────────────────────────────────

col1, col2, col3 = st.columns(3)

with col1:
    tipo_nf = st.selectbox("Tipo de NF", ["Entrada", "Saída"])

with col2:
    data_ini = st.date_input(
        "Data Inicial",
        value=datetime.now() - timedelta(days=30),
        format="YYYY-MM-DD",
    )

with col3:
    data_fim = st.date_input(
        "Data Final",
        value=datetime.now(),
        format="YYYY-MM-DD",
    )

# Filtro de loja (opcional)
lojas = buscar_lojas()
loja_opts = {l.get("nome", l.get("loja_id")): l.get("loja_id") for l in lojas}
loja_opts["Todas"] = None

col4, col5 = st.columns(2)
with col4:
    loja_selecionada = st.selectbox("Loja", list(loja_opts.keys()))
    loja_id = loja_opts[loja_selecionada]

with col5:
    st.write("")  # Espaçamento
    buscar_btn = st.button("🔍 Buscar", use_container_width=True, type="primary")

# ─────────────────────────────────────────────────────────────────────────────
# BUSCA E EXIBIÇÃO
# ─────────────────────────────────────────────────────────────────────────────

if buscar_btn:
    with st.spinner(f"Buscando NFs de {tipo_nf.lower()}..."):
        data_ini_str = data_ini.isoformat()
        data_fim_str = data_fim.isoformat()
        
        if tipo_nf == "Entrada":
            nfs = buscar_nfs_entrada(data_ini_str, data_fim_str, loja_id)
        else:
            nfs = buscar_nfs_saida(data_ini_str, data_fim_str, loja_id)
        
        if nfs:
            df = pd.DataFrame(nfs)
            
            # Formatar colunas monetárias
            cols_moeda = [
                "total_nf", "total_prod", "total_ipi", "total_frete",
                "total_icms", "total_st", "desconto"
            ]
            for col in cols_moeda:
                if col in df.columns:
                    df[col] = df[col].apply(lambda x: f"R$ {x:,.2f}")
            
            # Exibir estatísticas
            st.metric(f"Total de NFs de {tipo_nf}", len(df))
            
            # Exibir tabela
            st.dataframe(
                df,
                use_container_width=True,
                hide_index=True,
            )
            
            # Download CSV
            csv = df.to_csv(index=False, encoding="utf-8")
            st.download_button(
                label="📥 Baixar CSV",
                data=csv,
                file_name=f"nfs_{tipo_nf.lower()}_{data_ini_str}_{data_fim_str}.csv",
                mime="text/csv",
            )
        else:
            st.warning("Nenhuma NF encontrada no período.")

st.markdown("---")
st.caption(f"🔗 API URL: {API_URL}")
