"""
API FastAPI — Conferência NF de Entrada
Acessa Firebird e expõe endpoints REST para o Streamlit Cloud
Executa: python -m uvicorn conferencia_nf_api:app --host 0.0.0.0 --port 8000
"""

import os
import json
import fdb
from fastapi import FastAPI, HTTPException, Depends, Header
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from datetime import date
from typing import Optional, List
import logging

# ─── LOGGER ───────────────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ─── CONFIG ───────────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_CONFIG = os.path.join(BASE_DIR, 'db_config.json')
LOJAS_CONFIG = os.path.join(BASE_DIR, 'lojas_config.json')

# Token de autenticação (defina em variável de ambiente: API_TOKEN)
API_TOKEN = os.getenv("API_TOKEN", "seu_token_super_secreto_aqui")

FB_DLLS = [
    r'C:\Program Files (x86)\VisualControl\bin\fbclient25.dll',
    r'C:\Program Files\Firebird\Firebird_2_5\fbclient.dll',
    r'C:\Program Files (x86)\Firebird\Firebird_2_5\fbclient.dll',
    r'C:\Program Files\Firebird\Firebird_5_0\fbclient.dll',
    r'C:\Program Files\Firebird\Firebird_4_0\fbclient.dll',
    r'C:\Program Files\Firebird\Firebird_3_0\fbclient.dll',
]

# ─── INIT FB ───────────────────────────────────────────────────────────────────
for dll in FB_DLLS:
    if os.path.exists(dll):
        fdb.load_api(dll)
        logger.info(f"Firebird DLL carregada: {dll}")
        break
else:
    logger.warning("fbclient.dll não encontrada. Certifique-se de ter o Firebird Client instalado.")

# ─── MODELS ───────────────────────────────────────────────────────────────────
class NFEntrada(BaseModel):
    numero: str
    serie: str
    dt_entrada: date
    dt_emissao: date
    chave: Optional[str] = ""
    fornecedor: str
    cnpj_forn: str
    total_nf: float
    total_prod: float
    total_ipi: float
    total_frete: float
    total_icms: float
    total_st: float
    desconto: float
    loja: Optional[str] = ""

class NFSaida(BaseModel):
    numero: str
    serie: str
    dt_saida: date
    dt_emissao: date
    chave: Optional[str] = ""
    cliente: str
    cnpj_cli: str
    total_nf: float
    total_prod: float
    total_ipi: float
    total_frete: float
    total_icms: float
    total_st: float
    desconto: float
    loja: str
    xml: Optional[str] = ""

class HealthResponse(BaseModel):
    status: str
    message: str

# ─── HELPER FUNCTIONS ─────────────────────────────────────────────────────────
def _load_json(path, defaults):
    if os.path.exists(path):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                return {**defaults, **json.load(f)}
        except Exception as e:
            logger.error(f"Erro ao carregar {path}: {e}")
    return dict(defaults)

def load_db_config():
    return _load_json(DB_CONFIG, {
        'host': '192.168.15.240',
        'port': 3051,
        'database': r'C:\Program Files (x86)\VisualControl\Database\ERP_VCONTROL_BLINK.FDB',
        'user': 'SYSDBA',
        'password': 'masterkey',
    })

def load_lojas_config():
    """Carrega configuração de múltiplas lojas."""
    if not os.path.exists(LOJAS_CONFIG):
        return []
    try:
        with open(LOJAS_CONFIG, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        logger.error(f"Erro ao carregar {LOJAS_CONFIG}: {e}")
        return []

def _moeda(v):
    if v is None:
        return 0.0
    return float(v) if isinstance(v, (int, float, str)) else 0.0

def _limpar_cnpj(cnpj):
    if not cnpj:
        return ""
    s = str(cnpj).replace('.', '').replace('-', '').replace('/', '')
    return s

def conectar_erp(cfg):
    """Conecta ao banco Firebird."""
    try:
        return fdb.connect(
            host=cfg['host'],
            port=int(cfg['port']),
            database=cfg['database'],
            user=cfg['user'],
            password=cfg['password'],
            charset='UTF8',
        )
    except Exception as e:
        logger.error(f"Erro ao conectar Firebird: {e}")
        raise

def verify_token(authorization: Optional[str] = Header(None)):
    """Valida o token de autenticação."""
    if not authorization:
        raise HTTPException(status_code=401, detail="Falta token de autenticação")
    
    # Esperado formato: "Bearer TOKEN"
    parts = authorization.split()
    if len(parts) != 2 or parts[0].lower() != "bearer":
        raise HTTPException(status_code=401, detail="Formato de token inválido")
    
    token = parts[1]
    if token != API_TOKEN:
        raise HTTPException(status_code=403, detail="Token inválido")
    
    return token

# ─── FAST API APP ──────────────────────────────────────────────────────────────
app = FastAPI(
    title="Conferência NF — API",
    description="API para acesso aos dados de NF-e do Firebird",
    version="1.0.0",
)

# CORS para permitir Streamlit Cloud
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Mais seguro: lista específica de domínios
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ─── ENDPOINTS ─────────────────────────────────────────────────────────────────
@app.get("/api/health", response_model=HealthResponse)
async def health_check():
    """Verifica se a API está viva."""
    try:
        cfg = load_db_config()
        con = conectar_erp(cfg)
        con.close()
        return HealthResponse(status="OK", message="Firebird conectado com sucesso")
    except Exception as e:
        logger.error(f"Health check falhou: {e}")
        raise HTTPException(status_code=503, detail=f"Banco de dados indisponível: {str(e)}")

@app.get("/api/nfs-entrada", response_model=List[NFEntrada])
async def get_nfs_entrada(
    data_ini: str,  # Formato: YYYY-MM-DD
    data_fim: str,
    loja_id: Optional[int] = None,
    usar_emissao: bool = False,
    token: str = Depends(verify_token),
):
    """
    Busca notas fiscais de entrada.
    
    Args:
        data_ini: Data inicial (YYYY-MM-DD)
        data_fim: Data final (YYYY-MM-DD)
        loja_id: ID da loja (opcional)
        usar_emissao: Usar data de emissão ao invés de entrada
    """
    try:
        cfg = load_db_config()
        con = conectar_erp(cfg)
        cur = con.cursor()
        
        campo_data = 'e.EMISSAO' if usar_emissao else 'e.DATA_ENTRADA'
        filtro_loja = ''
        
        if loja_id is not None:
            prefixo = int(loja_id)
            id_min = prefixo * 10000000
            id_max = (prefixo + 1) * 10000000 - 1
            filtro_loja = f'AND e.ID BETWEEN {id_min} AND {id_max}'
        
        query = f"""
            SELECT
                e.NOTAFISCAL AS NUMERO,
                TRIM(e.SERIE) AS SERIE,
                e.DATA_ENTRADA AS DT_ENTRADA,
                e.EMISSAO AS DT_EMISSAO,
                '' AS CHAVE,
                TRIM(cl.NOME) AS FORNECEDOR,
                TRIM(cl.CPFCGC) AS CNPJ_FORN,
                CAST(e.TOTAL_NF AS DOUBLE PRECISION) AS TOTAL_NF,
                CAST(e.TOTAL_PROD AS DOUBLE PRECISION) AS TOTAL_PROD,
                CAST(COALESCE(e.TOTAL_IPI, 0) AS DOUBLE PRECISION) AS TOTAL_IPI,
                CAST(COALESCE(e.TOTAL_FRETE, 0) AS DOUBLE PRECISION) AS TOTAL_FRETE,
                CAST(COALESCE(e.VALOR_ICMS, 0) AS DOUBLE PRECISION) AS TOTAL_ICMS,
                CAST(COALESCE(e.VALOR_ICMS_SUBS, 0) AS DOUBLE PRECISION) AS TOTAL_ST,
                CAST(COALESCE(e.TOTAL_NF, 0) - COALESCE(e.TOTAL_LIQUIDO, e.TOTAL_NF, 0) AS DOUBLE PRECISION) AS DESCONTO,
                '' AS LOJA
            FROM ENTRADA_NF e
            INNER JOIN CLIFOR cl ON cl.ID = e.ID_FORNECEDOR
            WHERE COALESCE(e.STATUS, 'N') <> 'C'
              AND {campo_data} BETWEEN ? AND ?
              {filtro_loja}
            ORDER BY e.DATA_ENTRADA DESC, e.NOTAFISCAL
            ROWS 5000
        """
        
        cur.execute(query, (data_ini, data_fim))
        cols = [d[0] for d in cur.description]
        rows = []
        
        for row in cur.fetchall():
            r = dict(zip(cols, row))
            rows.append(NFEntrada(
                numero=str(r['NUMERO']),
                serie=(r['SERIE'] or "").strip(),
                dt_entrada=r['DT_ENTRADA'],
                dt_emissao=r['DT_EMISSAO'],
                chave=(r['CHAVE'] or "").strip(),
                fornecedor=r['FORNECEDOR'],
                cnpj_forn=_limpar_cnpj(r['CNPJ_FORN']),
                total_nf=_moeda(r['TOTAL_NF']),
                total_prod=_moeda(r['TOTAL_PROD']),
                total_ipi=_moeda(r['TOTAL_IPI']),
                total_frete=_moeda(r['TOTAL_FRETE']),
                total_icms=_moeda(r['TOTAL_ICMS']),
                total_st=_moeda(r['TOTAL_ST']),
                desconto=_moeda(r['DESCONTO']),
            ))
        
        cur.close()
        con.close()
        return rows
        
    except Exception as e:
        logger.error(f"Erro ao buscar NFs de entrada: {e}")
        raise HTTPException(status_code=500, detail=f"Erro na consulta: {str(e)}")

@app.get("/api/nfs-saida", response_model=List[NFSaida])
async def get_nfs_saida(
    data_ini: str,
    data_fim: str,
    loja_id: Optional[int] = None,
    token: str = Depends(verify_token),
):
    """
    Busca notas fiscais de saída (modelo 55).
    """
    try:
        cfg = load_db_config()
        con = conectar_erp(cfg)
        cur = con.cursor()
        
        params = [data_ini, data_fim]
        filtro_loja = ''
        
        if loja_id is not None:
            filtro_loja = 'AND nf.ID_LOJA = ?'
            params.append(int(loja_id))
        
        query = f"""
            SELECT
                nf.ID_COD AS NUMERO,
                TRIM(nf.SERIE) AS SERIE,
                nf.DATA AS DT_SAIDA,
                nf.EMISSAO AS DT_EMISSAO,
                nf.CHAVE_ACESSO_NFE AS CHAVE,
                TRIM(cl.NOME) AS CLIENTE,
                TRIM(cl.CPFCGC) AS CNPJ_CLI,
                CAST(nf.TOTAL_NF AS DOUBLE PRECISION) AS TOTAL_NF,
                CAST(nf.TOTAL_PROD AS DOUBLE PRECISION) AS TOTAL_PROD,
                CAST(nf.TOTAL_IPI AS DOUBLE PRECISION) AS TOTAL_IPI,
                CAST(nf.TOTAL_FRETE AS DOUBLE PRECISION) AS TOTAL_FRETE,
                CAST(nf.TOTAL_ICMS AS DOUBLE PRECISION) AS TOTAL_ICMS,
                CAST(nf.TOTAL_ICMS_ST AS DOUBLE PRECISION) AS TOTAL_ST,
                CAST(nf.DESCONTO_VALOR AS DOUBLE PRECISION) AS DESCONTO,
                TRIM(lj.NOME) AS LOJA,
                nf.XML AS XML_BLOB
            FROM NFISCAL nf
            INNER JOIN CLIFOR cl ON cl.ID = nf.ID_CLIFOR
            INNER JOIN LOJA lj ON lj.ID = nf.ID_LOJA
            WHERE nf.TIPO_ES = 'S'
              AND COALESCE(nf.STATUS, 'N') <> 'C'
              AND nf.DATA BETWEEN ? AND ?
              {filtro_loja}
            ORDER BY nf.DATA DESC, nf.ID_COD
            ROWS 5000
        """
        
        cur.execute(query, params)
        cols = [d[0] for d in cur.description]
        rows = []
        
        for row in cur.fetchall():
            r = dict(zip(cols, row))
            blob = r.pop('XML_BLOB', None)
            xml_str = ''
            if blob is not None:
                try:
                    raw = blob.read() if hasattr(blob, 'read') else blob
                    xml_str = raw.decode('utf-8', errors='ignore') if isinstance(raw, bytes) else (raw or '')
                except Exception:
                    pass
            
            rows.append(NFSaida(
                numero=str(r['NUMERO']),
                serie=(r['SERIE'] or "").strip(),
                dt_saida=r['DT_SAIDA'],
                dt_emissao=r['DT_EMISSAO'],
                chave=(r['CHAVE'] or "").strip(),
                cliente=r['CLIENTE'],
                cnpj_cli=_limpar_cnpj(r['CNPJ_CLI']),
                total_nf=_moeda(r['TOTAL_NF']),
                total_prod=_moeda(r['TOTAL_PROD']),
                total_ipi=_moeda(r['TOTAL_IPI']),
                total_frete=_moeda(r['TOTAL_FRETE']),
                total_icms=_moeda(r['TOTAL_ICMS']),
                total_st=_moeda(r['TOTAL_ST']),
                desconto=_moeda(r['DESCONTO']),
                loja=r['LOJA'],
                xml=xml_str,
            ))
        
        cur.close()
        con.close()
        return rows
        
    except Exception as e:
        logger.error(f"Erro ao buscar NFs de saída: {e}")
        raise HTTPException(status_code=500, detail=f"Erro na consulta: {str(e)}")

@app.get("/api/lojas")
async def get_lojas(token: str = Depends(verify_token)):
    """Lista todas as lojas configuradas."""
    try:
        lojas = load_lojas_config()
        return {"lojas": lojas}
    except Exception as e:
        logger.error(f"Erro ao listar lojas: {e}")
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        app,
        host="0.0.0.0",
        port=8000,
        log_level="info",
    )
