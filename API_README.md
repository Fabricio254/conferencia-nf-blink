# API FastAPI — Conferência NF

API REST para acesso aos dados de NF-e do Firebird. Permite que o Streamlit Cloud consume os dados sem precisar do `fdb` (que não funciona em Linux).

## 📋 Requisitos

- Python 3.11+
- Firebird Client instalado
- `requirements_api.txt`

## 🚀 Como Iniciar

### Local (Windows)

```bash
# Opção 1: Duplo clique em Iniciar_API.bat
Iniciar_API.bat

# Opção 2: Terminal
python -m uvicorn conferencia_nf_api:app --host 0.0.0.0 --port 8000 --reload
```

A API estará disponível em:
- 🌐 **HTTP**: http://localhost:8000
- 📚 **Swagger UI**: http://localhost:8000/docs
- 📖 **ReDoc**: http://localhost:8000/redoc

### Token de Autenticação

Defina a variável de ambiente `API_TOKEN`:

```bash
# Windows PowerShell
$env:API_TOKEN = "seu_token_super_secreto"
python -m uvicorn conferencia_nf_api:app --host 0.0.0.0 --port 8000

# Windows CMD
set API_TOKEN=seu_token_super_secreto
python -m uvicorn conferencia_nf_api:app --host 0.0.0.0 --port 8000
```

## 📡 Endpoints Disponíveis

### 1. Health Check
```bash
curl http://localhost:8000/api/health
```

**Resposta:**
```json
{
  "status": "OK",
  "message": "Firebird conectado com sucesso"
}
```

### 2. Buscar NFs de Entrada
```bash
curl -H "Authorization: Bearer seu_token" \
  "http://localhost:8000/api/nfs-entrada?data_ini=2025-05-01&data_fim=2025-05-31"
```

**Parâmetros:**
- `data_ini` (obrigatório): YYYY-MM-DD
- `data_fim` (obrigatório): YYYY-MM-DD
- `loja_id` (opcional): ID da loja
- `usar_emissao` (opcional): true/false

**Resposta:**
```json
[
  {
    "numero": "12345",
    "serie": "1",
    "dt_entrada": "2025-05-13",
    "dt_emissao": "2025-05-13",
    "chave": "12345678901234567890123456789012345678901234",
    "fornecedor": "FORNECEDOR LTDA",
    "cnpj_forn": "12345678000195",
    "total_nf": 10000.00,
    "total_prod": 8000.00,
    "total_ipi": 1000.00,
    "total_frete": 500.00,
    "total_icms": 400.00,
    "total_st": 100.00,
    "desconto": 0.00,
    "loja": ""
  }
]
```

### 3. Buscar NFs de Saída
```bash
curl -H "Authorization: Bearer seu_token" \
  "http://localhost:8000/api/nfs-saida?data_ini=2025-05-01&data_fim=2025-05-31"
```

**Parâmetros:**
- `data_ini` (obrigatório): YYYY-MM-DD
- `data_fim` (obrigatório): YYYY-MM-DD
- `loja_id` (opcional): ID da loja

### 4. Listar Lojas
```bash
curl -H "Authorization: Bearer seu_token" \
  "http://localhost:8000/api/lojas"
```

## 🔌 Consumir no Streamlit (Exemplo)

```python
import requests
import streamlit as st
from datetime import date

# Configuração
API_URL = "http://seu-servidor.com:8000"
API_TOKEN = st.secrets.get("api_token", "seu_token")

# Headers com autenticação
headers = {"Authorization": f"Bearer {API_TOKEN}"}

# Buscar NFs
data_ini = st.date_input("Data inicial")
data_fim = st.date_input("Data final")

if st.button("Buscar"):
    try:
        response = requests.get(
            f"{API_URL}/api/nfs-entrada",
            params={
                "data_ini": data_ini.isoformat(),
                "data_fim": data_fim.isoformat(),
            },
            headers=headers,
            timeout=30,
        )
        response.raise_for_status()
        
        nfs = response.json()
        st.dataframe(nfs)
        
    except requests.exceptions.RequestException as e:
        st.error(f"Erro na API: {e}")
```

## 🌐 Deployment no Streamlit Cloud

1. **Adicionar `.streamlit/secrets.toml` no Streamlit Cloud:**
```toml
api_url = "https://seu-servidor.com:8000"
api_token = "seu_token_super_secreto"
```

2. **Usar no Streamlit:**
```python
import requests
import streamlit as st

API_URL = st.secrets["api_url"]
API_TOKEN = st.secrets["api_token"]

response = requests.get(
    f"{API_URL}/api/nfs-entrada",
    params={"data_ini": "2025-05-01", "data_fim": "2025-05-31"},
    headers={"Authorization": f"Bearer {API_TOKEN}"},
)
```

## 🔒 Segurança

### Para Produção

1. **Use HTTPS** com certificado SSL
```bash
python -m uvicorn conferencia_nf_api:app \
  --host 0.0.0.0 \
  --port 8443 \
  --ssl-keyfile=/path/to/key.pem \
  --ssl-certfile=/path/to/cert.pem
```

2. **Use ngrok** para expor a API com URL segura
```bash
ngrok http 8000 --domain=seu-dominio-fixo.ngrok-free.dev
```

3. **Altere o token padrão**
```bash
set API_TOKEN=um_token_aleatorio_muito_complexo
```

4. **Restrinja CORS** editando `conferencia_nf_api.py`:
```python
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://conferencia-nf-blink.streamlit.app"],
    allow_credentials=True,
    allow_methods=["GET", "POST"],
    allow_headers=["Authorization"],
)
```

## 📊 Arquitetura

```
┌─────────────────────┐
│ Streamlit Cloud     │
│ (conferencia-nf...) │
└──────────┬──────────┘
           │ requests HTTP
           │ (Bearer token)
           ↓
┌─────────────────────┐
│ API FastAPI         │
│ (sua-máquina:8000)  │
└──────────┬──────────┘
           │ fdb
           ↓
┌─────────────────────┐
│ Firebird            │
│ (192.168.15.240)    │
└─────────────────────┘
```

## ❓ FAQ

**P: A API pode ficar offline?**  
R: Sim. Para produção, configure:
- Systemd/Windows Service para restart automático
- Monitoramento com alerts
- Failover redundante

**P: Como aumentar o timeout?**  
R: Edite `conferencia_nf_api.py` e altere `ROWS 5000` para menos registros ou configure índices no Firebird.

**P: Como autenticar sem token?**  
R: Remova a função `verify_token` dos decoradores `@app.get()`. ⚠️ Não recomendado em produção.

## 📝 Logs

Os logs são exibidos no console. Para salvar em arquivo:
```bash
python -m uvicorn conferencia_nf_api:app --host 0.0.0.0 --port 8000 > api.log 2>&1
```

## 🐛 Troubleshooting

**"fbclient.dll não encontrada"**
- Instale o Firebird Client: https://www.firebirdsql.org/

**"Connection refused"**
- Verifique se `db_config.json` tem o host/port corretos
- Verifique se o Firebird está rodando
- Teste com: `telnet 192.168.15.240 3051`

**"Authentication failed"**
- Verifique usuário/senha em `db_config.json`
- Consulte logs do Firebird

## 📞 Suporte

Para relatórios de bugs ou sugestões, entre em contato com o time de TI.
