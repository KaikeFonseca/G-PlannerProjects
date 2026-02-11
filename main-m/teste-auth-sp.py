# -*- coding: utf-8 -*-
import os
import sys
import json
import html
import requests
import pandas as pd
from datetime import datetime, date
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.listitems.caml.query import CamlQuery
# ====================== CONFIG ======================
SHAREPOINT_URL = 'https://gitservices.sharepoint.com/sites/LogsticaAA3'  # sem barra final
TENANT_HOST    = 'gitservices.sharepoint.com'

# Pegue as credenciais pelas variáveis de ambiente
USERNAME = os.getenv("SPO_USER")
PASSWORD = os.getenv("SPO_PASS")

# Nomes de listas (ajuste conforme o seu tenant)
PLANNER_LIST_NAME            = "list_automacaoPlanner_py"
PLANNER_LIST_NAME_RECEBER    = "list_solicitacao_planner_py"

# Mapeamentos (mantive os seus)
DF_TO_SP_MAP = {
    'Material': 'material',
    'posto': 'posto',
    'tags': 'tags',
    'data': 'data',
    'linha': 'linha',
    'patan': 'patan',
    'turno': 'turno',
    'tempPeca': 'tempo_por_peca',
    'qtdPecasSeremProduzidas': 'qtd_pecas_serem_produzidas',
    'qtdPorKanban': 'qtd_por_kanban',
    'kanbans': 'kanbans',
    'tempoProd': 'tempo_prod',
    'sequencia': 'sequencia',
    'prodEmLinha': 'prod_em_linha',
    'compComb': 'comp_comb',
    'estoqueMaterial': 'estoque_material_atual',
    'estoqueKanbanMax': 'estoque_kanban_maximo',
    'diff': 'diff_estoque_atual_e_estoque_max',
    'obs': 'obs',
    'STATUS': 'STATUS',
    'horaProdInicial': 'hora_prod_inicial',
    'horaProdFinal': 'hora_prod_final',
    'descricaoRefeicao': 'descricao_refeicao',
    'checklist': 'checklist',
    'descricao': 'descricao_planner'
}

RECEIVE_SP_MAP = {
    'ID':    'ID',
    'linha': 'linha',
    'patan': 'patan',
    'turno': 'turno',
    'STATUS':'STATUS'
}
# ====================================================


# =============== TESTE A: usando a lib office365 ===============
def get_sharepoint_context_office365():
    """Autentica de forma legada com a lib office365-rest-python-client."""
    if not USERNAME or not PASSWORD:
        raise RuntimeError("Defina as variáveis de ambiente SPO_USER e SPO_PASS.")

    # A lib usa AuthenticationContext + acquire_token_for_user (legado) -> wsignin1.0 por baixo dos panos
    auth_ctx = AuthenticationContext(SHAREPOINT_URL)
    auth_ctx.acquire_token_for_user(username=USERNAME, password=PASSWORD)
    ctx = ClientContext(SHAREPOINT_URL, auth_ctx)
    return ctx

def test_office365_read_web_title():
    """Lê o título do site (sanity check da autenticação)."""
    print("\n[Teste A1] Autenticando (office365 lib) e lendo o título do site...")
    ctx = get_sharepoint_context_office365()
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()
    print(f"   ✓ Título do site: {web.properties.get('Title')}")

def test_office365_list_top1(list_name: str):
    """Busca 1 item da lista de recebimento com STATUS = 1, via CAML."""
    print(f"\n[Teste A2] Lendo 1 item de '{list_name}' com STATUS=1 (CAML)...")
    ctx = get_sharepoint_context_office365()
    sp_list = ctx.web.lists.get_by_title(list_name)

    # CAML simples: ajuste o Type conforme o tipo da coluna STATUS
    where_clause = "<Where><Eq><FieldRef Name='STATUS'/><Value Type='Text'>1</Value></Eq></Where>"
    view_fields_xml = "".join([f"<FieldRef Name='{f}'/>" for f in RECEIVE_SP_MAP.keys()])

    caml = CamlQuery()
    caml.ViewXml = f"<View><RowLimit>1</RowLimit><ViewFields>{view_fields_xml}</ViewFields><Query>{where_clause}</Query></View>"

    items = sp_list.get_items(caml).execute_query()
    if not items:
        print("   ⚠ Nenhum item retornado (STATUS=1). Verifique se a coluna STATUS é Texto ou Número.")
        return

    item = items[0]
    dados = {df_col: item.properties.get(sp_col) for (sp_col, df_col) in RECEIVE_SP_MAP.items()}
    print("   ✓ Item de exemplo:", dados)


# =============== TESTE B: raw legacy (SAML -> wsignin1.0) ===============
class LegacyAuthError(Exception):
    pass

def get_saml_bst(tenant_host: str, username: str, password: str) -> str:
    """Obtém BinarySecurityToken no extSTS.srf (precisa ser gerado a cada login)."""
    soap = f"""<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope"
    xmlns:a="http://www.w3.org/2005/08/addressing"
    xmlns:u="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
    <s:Header>
        <a:Action s:mustUnderstand="1">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</a:Action>
        <a:ReplyTo><a:Address>http://www.w3.org/2005/08/addressing/anonymous</a:Address></a:ReplyTo>
        <a:To s:mustUnderstand="1">https://login.microsoftonline.com/extSTS.srf</a:To>
        <o:Security s:mustUnderstand="1"
        xmlns:o="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
        <o:UsernameToken>
            <o:Username>{html.escape(username)}</o:Username>
            <o:Password>{html.escape(password)}</o:Password>
        </o:UsernameToken>
        </o:Security>
    </s:Header>
    <s:Body>
        <t:RequestSecurityToken xmlns:t="http://schemas.xmlsoap.org/ws/2005/02/trust">
        <wsp:AppliesTo xmlns:wsp="http://schemas.xmlsoap.org/ws/2004/09/policy">
            <a:EndpointReference><a:Address>https://{tenant_host}/</a:Address></a:EndpointReference>
        </wsp:AppliesTo>
        <t:KeyType>http://schemas.xmlsoap.org/ws/2005/05/identity/NoProofKey</t:KeyType>
        <t:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</t:RequestType>
        <t:TokenType>urn:oasis:names:tc:SAML:1.0:assertion</t:TokenType>
        </t:RequestSecurityToken>
    </s:Body>
    </s:Envelope>"""
    headers = {"Content-Type": "application/soap+xml; charset=utf-8","User-Agent": "PythonRequestsLegacyWSFed/1.0"}
    r = requests.post("https://login.microsoftonline.com/extSTS.srf",data=soap.encode("utf-8"), headers=headers, timeout=30)
    r.raise_for_status()

    # parse simples para achar o BinarySecurityToken
    import xml.etree.ElementTree as ET
    root = ET.fromstring(r.text)
    for elem in root.iter():
        if elem.tag.endswith("BinarySecurityToken"):
            return elem.text
    raise LegacyAuthError("BinarySecurityToken não encontrado (possível MFA/CA ou domínio federado).")

def legacy_login_session(tenant_host: str, username: str, password: str) -> requests.Session:
    """POST no wsignin1.0 e retorna sessão com cookies FedAuth/rtFa (se liberado no tenant)."""
    bst = get_saml_bst(tenant_host, username, password)
    s = requests.Session()
    s.headers.update({"User-Agent": "PythonRequestsLegacyWSFed/1.0"})
    resp = s.post(f"https://{tenant_host}/_forms/default.aspx?wa=wsignin1.0",data={"wa": "wsignin1.0", "wresult": bst},allow_redirects=True, timeout=30)
    print(f"[Teste B] POST wsignin1.0 -> HTTP {resp.status_code}")
    print("          Cookies:", list(s.cookies.get_dict().keys()))
    if not any(k.lower() in ("fedauth", "rtfa") for k in s.cookies.get_dict().keys()):
        raise LegacyAuthError("Sem cookies FedAuth/rtFa. Provável bloqueio de autenticação legada ou exigência de MFA.")
    return s

def test_raw_legacy_read_web():
    """Com a sessão de cookies, chama _api/web para confirmar acesso."""
    if not USERNAME or not PASSWORD:
        raise RuntimeError("Defina as variáveis SPO_USER e SPO_PASS.")
    s = legacy_login_session(TENANT_HOST, USERNAME, PASSWORD)
    headers = {"Accept": "application/json;odata=nometadata"}
    r = s.get(f"{SHAREPOINT_URL}/_api/web?$select=Title,Url,Id", headers=headers, timeout=30)
    print("[Teste B] GET _api/web ->", r.status_code, r.text[:300])


# =========================== MAIN ===========================
if __name__ == "__main__":
    print("== Teste de Autenticação e Requisições ao SharePoint ==")
    if not USERNAME or not PASSWORD:
        print("ERRO: defina as variáveis de ambiente SPO_USER e SPO_PASS.")
        sys.exit(1)

    try:
        # ---------- Teste A: biblioteca office365 ----------
        test_office365_read_web_title()
        test_office365_list_top1(PLANNER_LIST_NAME_RECEBER)

    except Exception as e:
        print("\n[Falha no Teste A (office365 lib)]")
        print("Detalhes:", e)

    try:
        # ---------- Teste B: fluxo raw legado ----------
        test_raw_legacy_read_web()

    except Exception as e:
        print("\n[Falha no Teste B (raw legado))]")
        print("Detalhes:", e)
        print("\nSe for HTTP 403 / sem FedAuth/rtFa, normalmente é bloqueio de 'Apps que não usam autenticação moderna' "
                "ou política de Acesso Condicional/MFA. Peça liberação temporária ao TI.")
