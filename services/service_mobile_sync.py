from __future__ import annotations

import json
import os
import shutil
import socket
import threading
from datetime import datetime
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs

from utils.caminhos import caminho_mobile_importados, caminho_mobile_pendencias

HOST_PADRAO = "0.0.0.0"
PORTA_PADRAO = 8765

_SERVIDOR = None
_THREAD = None


HTML_FORM = """<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Coleta Mobile</title>
  <style>
    :root {{
      --bg: #f3f6fb;
      --panel: rgba(255,255,255,.92);
      --card: #ffffff;
      --card-soft: #f8fbff;
      --border: #d6e0ee;
      --text: #0f172a;
      --muted: #61738f;
      --accent: #d96b1d;
      --accent-dark: #b7540e;
      --secondary: #1d3557;
      --secondary-soft: #edf3fb;
      --shadow: 0 20px 48px rgba(15, 23, 42, .10);
      --radius-xl: 28px;
      --radius-lg: 22px;
      --radius-md: 18px;
      --radius-sm: 14px;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "Segoe UI", sans-serif;
      background:
        radial-gradient(circle at top left, rgba(217,107,29,.18), transparent 28%),
        linear-gradient(180deg, #f8fbff 0%, #eaf2ff 100%);
      color: var(--text);
    }}
    .wrap {{ max-width: 820px; margin: 0 auto; padding: 18px 14px 132px; }}
    .hero {{
      background: linear-gradient(135deg, #10213f 0%, #1a2f56 100%);
      color: #fff;
      border-radius: var(--radius-xl);
      padding: 24px 20px;
      box-shadow: var(--shadow);
      margin-bottom: 18px;
      position: relative;
      overflow: hidden;
    }}
    .hero::after {{
      content: "";
      position: absolute;
      width: 180px;
      height: 180px;
      right: -70px;
      top: -90px;
      background: radial-gradient(circle, rgba(217,107,29,.24), transparent 70%);
    }}
    .hero-top {{
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      gap: 12px;
      margin-bottom: 14px;
    }}
    .badge {{
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 8px 12px;
      border-radius: 999px;
      background: rgba(255,255,255,.12);
      font-size: 12px;
      font-weight: 700;
      color: #d9e6ff;
      margin-bottom: 12px;
    }}
    h1 {{ margin: 0 0 10px; font-size: 32px; line-height: 1.02; max-width: 12ch; }}
    p {{ color: var(--muted); }}
    .hero p {{ color: #d5def1; margin: 0; font-size: 14px; line-height: 1.55; max-width: 46ch; }}
    .summary {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 12px;
      margin-top: 16px;
    }}
    .sum-card {{
      background: rgba(255,255,255,.08);
      border: 1px solid rgba(255,255,255,.1);
      border-radius: 18px;
      padding: 14px;
      backdrop-filter: blur(8px);
    }}
    .sum-card strong {{
      display: block;
      font-size: 20px;
      margin-bottom: 5px;
    }}
    .sum-card span {{ font-size: 12px; color: #d7e1f3; }}
    .card {{
      background: var(--panel);
      border: 1px solid rgba(255,255,255,.8);
      backdrop-filter: blur(14px);
      border-radius: var(--radius-lg);
      padding: 20px;
      box-shadow: var(--shadow);
    }}
    .lead {{
      margin: 0 0 18px;
      color: var(--muted);
      font-size: 14px;
      line-height: 1.6;
    }}
    .grid {{ display: grid; gap: 16px; grid-template-columns: repeat(2, minmax(0, 1fr)); }}
    .full {{ grid-column: 1 / -1; }}
    label {{
      display: block;
      font-size: 11px;
      font-weight: 700;
      letter-spacing: .06em;
      margin-bottom: 8px;
      color: var(--secondary);
      text-transform: uppercase;
    }}
    input, select, textarea, button {{
      width: 100%;
      border: 1px solid var(--border);
      border-radius: 16px;
      padding: 15px 16px;
      background: var(--card);
      color: var(--text);
      font: inherit;
      font-size: 16px;
      line-height: 1.3;
      box-shadow: none;
    }}
    input:focus, select:focus, textarea:focus {{
      outline: none;
      border-color: #7da1cf;
      box-shadow: 0 0 0 3px rgba(125, 161, 207, .18);
    }}
    input::placeholder, textarea::placeholder {{ color: #8fa0b8; }}
    textarea {{ min-height: 136px; resize: vertical; }}
    button {{
      border: none;
      font-weight: 700;
      margin-top: 0;
      transition: transform .12s ease, background .12s ease;
      cursor: pointer;
    }}
    button:active {{ transform: translateY(1px); }}
    .hint {{
      font-size: 12px;
      color: var(--muted);
      margin-top: 8px;
      line-height: 1.5;
      background: #f4f8fd;
      border: 1px solid #dce6f2;
      border-radius: 14px;
      padding: 10px 12px;
    }}
    .item {{
      border: 1px solid #d7e2f1;
      border-radius: 24px;
      padding: 18px;
      margin-top: 16px;
      background: linear-gradient(180deg, #ffffff 0%, var(--card-soft) 100%);
      box-shadow: 0 10px 24px rgba(15, 23, 42, .06);
    }}
    .item-top {{
      display:flex;
      justify-content:space-between;
      align-items:center;
      gap:12px;
      margin-bottom:16px;
    }}
    .item-title {{
      font-weight: 800;
      font-size: 19px;
      color: var(--secondary);
    }}
    .item-sub {{
      font-size: 13px;
      color: var(--muted);
      margin-top: 4px;
      line-height: 1.45;
    }}
    .pill {{
      display:inline-flex;
      align-items:center;
      justify-content:center;
      min-width: 38px;
      height: 38px;
      padding: 0 13px;
      border-radius: 999px;
      background: #eef3fa;
      color: var(--secondary);
      font-size: 12px;
      font-weight: 800;
      border: 1px solid #d7e2f1;
    }}
    .sec {{
      background: var(--secondary);
      color: #fff;
      border: none;
      padding: 12px 16px;
      border-radius: 14px;
      width: auto;
      font-size: 14px;
    }}
    .ghost {{ background: #e8eef7; color: var(--secondary); }}
    .actions {{
      position: fixed;
      left: 0;
      right: 0;
      bottom: 0;
      padding: 14px 14px 18px;
      background: linear-gradient(180deg, rgba(248,251,255,0) 0%, rgba(248,251,255,.96) 30%, rgba(248,251,255,1) 100%);
      backdrop-filter: blur(8px);
    }}
    .actions-inner {{
      max-width: 820px;
      margin: 0 auto;
      display: grid;
      gap: 12px;
    }}
    .primary {{ background: var(--accent); color: #fff; }}
    .primary:hover {{ background: var(--accent-dark); }}
    .secondary-btn {{ background: #ffffff; color: var(--secondary); border: 1px solid #d7e2f1; }}
    .section-chip {{
      display: inline-flex;
      align-items: center;
      gap: 8px;
      background: #eff4fb;
      border: 1px solid #dce6f2;
      border-radius: 999px;
      padding: 7px 12px;
      font-size: 12px;
      color: var(--secondary);
      font-weight: 700;
      margin-bottom: 14px;
    }}
    .field {{
      display: flex;
      flex-direction: column;
    }}
    .field-soft {{
      background: #f9fbfe;
      border: 1px solid #e3ebf6;
      border-radius: 18px;
      padding: 12px;
    }}
    .divider {{
      height: 1px;
      background: linear-gradient(90deg, transparent, #dce6f2, transparent);
      margin: 16px 0 2px;
    }}
    @media (max-width: 640px) {{
      .grid, .summary {{ grid-template-columns: 1fr; }}
      .item-top {{ align-items: flex-start; flex-direction: column; }}
      .wrap {{ padding-bottom: 148px; }}
      .card {{ padding: 18px 16px; }}
      .item {{ padding: 16px; }}
      .actions {{ padding-left: 12px; padding-right: 12px; }}
    }}
  </style>
</head>
<body>
  <div class="wrap">
    <div class="hero">
      <div class="hero-top">
        <div>
          <div class="badge">Coleta Mobile</div>
          <h1>Envio de manutencoes</h1>
          <p>Preencha um ou varios formularios, tire as fotos no celular e envie tudo para o sistema desktop.</p>
        </div>
      </div>
      <div class="summary">
        <div class="sum-card">
          <strong id="total-formularios">1</strong>
          <span>Formulario(s) no lote</span>
        </div>
        <div class="sum-card">
          <strong>Wi-Fi</strong>
          <span>Envio direto para o computador</span>
        </div>
      </div>
    </div>
    <div class="card">
      <p class="lead">Cada formulario pode representar um veiculo, equipamento ou manutencao diferente no mesmo envio.</p>
      <form id="mobile-form" method="post" action="/upload" enctype="multipart/form-data">
        <input type="hidden" name="item_count" id="item_count" value="1">
        <div id="items"></div>
      </form>
    </div>
  </div>
  <div class="actions">
    <div class="actions-inner">
      <button type="button" class="secondary-btn" onclick="adicionarItem()">Adicionar outro formulario</button>
      <button type="submit" form="mobile-form" class="primary">Enviar lote para o sistema</button>
    </div>
  </div>
  <template id="item-template">
    <div class="item" data-index="__INDEX__">
      <div class="item-top">
        <div>
          <div class="item-title">Formulario __NUM__</div>
          <div class="item-sub">Preencha os dados e anexe as fotos deste item.</div>
        </div>
        <div style="display:flex;gap:8px;align-items:center">
          <span class="pill">#__NUM__</span>
          <button type="button" class="sec ghost" onclick="removerItem(this)">Remover</button>
        </div>
      </div>
      <div class="section-chip">Dados basicos</div>
      <div class="grid">
        <div class="field">
          <label>Destino do envio</label>
          <select name="destino____INDEX__">
            <option value="MANUTENCAO">Manutencao</option>
            <option value="VEICULO">Veiculo cadastrado</option>
          </select>
        </div>
        <div class="field">
          <label>Patrimonio</label>
          <input name="patrimonio____INDEX__" required placeholder="Ex.: ABC1234">
        </div>
        <div class="field">
          <label>Categoria</label>
          <select name="categoria____INDEX__">
            <option>GERAL</option>
            <option>OLEO</option>
            <option>PNEUS</option>
            <option>MECANICA</option>
            <option>ELETRICA</option>
          </select>
        </div>
        <div class="full field">
          <label>Descricao</label>
          <input name="descricao____INDEX__" required placeholder="Descreva a manutencao">
        </div>
        <div class="field">
          <label>Marca</label>
          <input name="marca____INDEX__" placeholder="Ex.: Volvo, Caterpillar">
        </div>
        <div class="field">
          <label>Ano</label>
          <input name="ano____INDEX__" inputmode="numeric" placeholder="Ex.: 2024">
        </div>
      </div>
      <div class="divider"></div>
      <div class="section-chip">Medicoes e datas</div>
      <div class="grid">
        <div class="field">
          <label>Valor</label>
          <input name="valor____INDEX__" inputmode="decimal" placeholder="0,00">
        </div>
        <div class="field">
          <label>Horimetro atual</label>
          <input name="horimetro_atual____INDEX__" inputmode="decimal" placeholder="Ex.: 1250">
        </div>
        <div class="field">
          <label>Horimetro troca</label>
          <input name="horimetro_troca____INDEX__" inputmode="decimal" placeholder="Ex.: 1500">
        </div>
        <div class="field">
          <label>Data inicio</label>
          <input name="data_inicio____INDEX__" placeholder="dd/mm/aaaa">
        </div>
        <div class="field">
          <label>Data fim</label>
          <input name="data_fim____INDEX__" placeholder="dd/mm/aaaa">
        </div>
      </div>
      <div class="divider"></div>
      <div class="section-chip">Detalhes e fotos</div>
      <div class="grid">
        <div class="full field">
          <label>Detalhe</label>
          <textarea name="detalhe____INDEX__" placeholder="Observacoes da manutencao"></textarea>
        </div>
        <div class="full field field-soft">
          <label>Fotos</label>
          <input type="file" name="fotos____INDEX__" accept="image/*" multiple>
          <div class="hint">Selecione varias imagens da galeria ou da camera. Em muitos celulares, remover o modo direto da camera melhora o envio multiplo.</div>
        </div>
      </div>
    </div>
  </template>
  <script>
    let contador = 0;
    function render() {
      document.querySelectorAll('.item').forEach((el, i) => {
        el.querySelector('.item-title').textContent = `Formulario ${i + 1}`;
        el.querySelector('.item-sub').textContent = `Preencha os dados e anexe as fotos deste item.`;
        el.querySelector('.pill').textContent = `#${i + 1}`;
      });
      const total = document.querySelectorAll('.item').length;
      document.getElementById('item_count').value = total;
      document.getElementById('total-formularios').textContent = total;
    }
    function adicionarItem() {
      const tpl = document.getElementById('item-template').innerHTML
        .replaceAll('__INDEX__', String(contador))
        .replaceAll('__NUM__', String(contador + 1));
      document.getElementById('items').insertAdjacentHTML('beforeend', tpl);
      contador += 1;
      render();
    }
    function removerItem(botao) {
      const items = document.querySelectorAll('.item');
      if (items.length <= 1) return;
      botao.closest('.item').remove();
      render();
    }
    adicionarItem();
  </script>
</body>
</html>
"""


def _slug(texto: str) -> str:
    permitido = "".join(ch for ch in str(texto) if ch.isalnum() or ch in ("-", "_", "."))
    return permitido.strip("._-") or "arquivo"


def obter_ip_local() -> str:
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
            sock.connect(("8.8.8.8", 80))
            return sock.getsockname()[0]
    except Exception:
        try:
            return socket.gethostbyname(socket.gethostname())
        except Exception:
            return "127.0.0.1"


def obter_url_mobile(porta: int = PORTA_PADRAO) -> str:
    return f"http://{obter_ip_local()}:{porta}"


def _organizar_itens(campos: dict[str, str], arquivos: list[tuple[str, bytes]]) -> tuple[list[dict], list[str]]:
    itens: dict[str, dict] = {}
    fotos_livres: list[str] = []

    for chave, valor in campos.items():
        if "__" in chave:
            nome, indice = chave.rsplit("__", 1)
            item = itens.setdefault(indice, {"dados": {}, "fotos": []})
            item["dados"][nome] = valor

    for campo_nome, nome_arquivo, conteudo in arquivos:
        if "__" in campo_nome:
            campo, indice = campo_nome.rsplit("__", 1)
            if campo == "fotos":
                item = itens.setdefault(indice, {"dados": {}, "fotos": []})
                item["fotos"].append((campo_nome, nome_arquivo, conteudo))
            else:
                fotos_livres.append((campo_nome, nome_arquivo, conteudo))
        else:
            fotos_livres.append((campo_nome, nome_arquivo, conteudo))

    lista = []
    for indice in sorted(itens.keys(), key=lambda valor: int(valor) if str(valor).isdigit() else str(valor)):
        lista.append(itens[indice])
    return lista, fotos_livres


def _salvar_payload(campos: dict[str, str], arquivos: list[tuple[str, str, bytes]], origem: str | None = None) -> str:
    submission_id = datetime.now().strftime("%Y%m%d%H%M%S%f")
    pasta = Path(caminho_mobile_pendencias()) / submission_id
    pasta.mkdir(parents=True, exist_ok=True)

    itens, fotos_livres = _organizar_itens(campos, arquivos)

    def salvar_arquivo(nome_original: str, conteudo: bytes, subpasta: Path) -> str:
        subpasta.mkdir(parents=True, exist_ok=True)
        nome_limpo = _slug(nome_original)
        destino = subpasta / nome_limpo
        contador = 1
        while destino.exists():
            destino = subpasta / f"{Path(nome_limpo).stem}_{contador}{Path(nome_limpo).suffix}"
            contador += 1
        destino.write_bytes(conteudo)
        return str(destino.relative_to(pasta))

    itens_salvos = []
    for indice, item in enumerate(itens):
        pasta_item = pasta / f"item_{indice+1:02d}"
        fotos_item = [salvar_arquivo(nome, conteudo, pasta_item / "fotos") for _campo, nome, conteudo in item["fotos"]]
        itens_salvos.append(
            {
                "ordem": indice + 1,
                "dados": item["dados"],
                "fotos": fotos_item,
            }
        )

    anexos_gerais = [salvar_arquivo(nome, conteudo, pasta / "anexos_gerais") for _campo, nome, conteudo in fotos_livres]

    payload = {
        "id": submission_id,
        "origem": origem or "",
        "recebido_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "itens": itens_salvos,
        "anexos_gerais": anexos_gerais,
        "status": "pendente",
    }
    (pasta / "manifest.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return submission_id


def listar_pendencias_mobile() -> list[dict]:
    base = Path(caminho_mobile_pendencias())
    resultados = []
    for pasta in sorted(base.iterdir(), reverse=True):
        manifest = pasta / "manifest.json"
        if not manifest.exists():
            continue
        try:
            payload = json.loads(manifest.read_text(encoding="utf-8"))
        except Exception:
            continue
        payload["pasta"] = str(pasta)
        resultados.append(payload)
    return resultados


def obter_pendencia_mobile(submission_id: str) -> dict | None:
    pasta = Path(caminho_mobile_pendencias()) / submission_id
    manifest = pasta / "manifest.json"
    if not manifest.exists():
        return None
    payload = json.loads(manifest.read_text(encoding="utf-8"))
    payload["pasta"] = str(pasta)
    return payload


def concluir_pendencia_mobile(submission_id: str) -> None:
    origem = Path(caminho_mobile_pendencias()) / submission_id
    destino = Path(caminho_mobile_importados()) / submission_id
    if not origem.exists():
        return
    if destino.exists():
        shutil.rmtree(destino)
    shutil.move(str(origem), str(destino))


def excluir_pendencia_mobile(submission_id: str) -> None:
    pasta = Path(caminho_mobile_pendencias()) / submission_id
    if pasta.exists():
        shutil.rmtree(pasta)


def _parse_request(handler: BaseHTTPRequestHandler) -> tuple[dict[str, str], list[tuple[str, str, bytes]]]:
    content_type = handler.headers.get("Content-Type", "")
    content_length = int(handler.headers.get("Content-Length", "0") or 0)
    body = handler.rfile.read(content_length)

    if "multipart/form-data" in content_type:
        header = f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8")
        mensagem = BytesParser(policy=default).parsebytes(header + body)
        campos: dict[str, str] = {}
        arquivos: list[tuple[str, str, bytes]] = []
        for parte in mensagem.iter_parts():
            name = parte.get_param("name", header="content-disposition")
            filename = parte.get_filename()
            conteudo = parte.get_payload(decode=True) or b""
            if filename:
                arquivos.append((name or "arquivo", filename, conteudo))
            elif name:
                campos[name] = conteudo.decode("utf-8", errors="ignore").strip()
        return campos, arquivos

    if "application/x-www-form-urlencoded" in content_type:
        parsed = parse_qs(body.decode("utf-8", errors="ignore"))
        return {chave: valores[0] for chave, valores in parsed.items()}, []

    return {}, []


class _MobileHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path not in ("/", "/index.html"):
            self.send_error(404)
            return
        html = HTML_FORM.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(html)))
        self.end_headers()
        self.wfile.write(html)

    def do_POST(self):
        if self.path != "/upload":
            self.send_error(404)
            return

        campos, arquivos = _parse_request(self)
        indices_validos = sorted(
            {
                chave.rsplit("__", 1)[1]
                for chave in campos
                if "__" in chave and campos.get(chave, "").strip() and chave.startswith(("patrimonio__", "descricao__"))
            }
        )
        if not indices_validos:
            self.send_error(400, "Informe ao menos um formulario com patrimonio e descricao.")
            return

        submission_id = _salvar_payload(campos, arquivos, origem=self.client_address[0])
        html = f"""<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"></head>
<body style="font-family:Segoe UI,sans-serif;background:#0f172a;color:#f8fafc;padding:24px">
<h2>Dados enviados com sucesso</h2>
<p>ID da coleta: {submission_id}</p>
<p>Lote recebido. Volte ao sistema desktop e use a tela de Coleta Mobile para importar.</p>
<a href="/" style="color:#f97316">Enviar nova coleta</a>
</body></html>""".encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(html)))
        self.end_headers()
        self.wfile.write(html)

    def log_message(self, *_args):
        return


def iniciar_servidor_mobile(porta: int = PORTA_PADRAO) -> str:
    global _SERVIDOR, _THREAD
    if _SERVIDOR is not None:
        return obter_url_mobile(porta)

    _SERVIDOR = ThreadingHTTPServer((HOST_PADRAO, porta), _MobileHandler)
    _THREAD = threading.Thread(target=_SERVIDOR.serve_forever, daemon=True)
    _THREAD.start()
    return obter_url_mobile(porta)


def parar_servidor_mobile() -> None:
    global _SERVIDOR, _THREAD
    if _SERVIDOR is None:
        return
    _SERVIDOR.shutdown()
    _SERVIDOR.server_close()
    _SERVIDOR = None
    _THREAD = None


def servidor_mobile_ativo() -> bool:
    return _SERVIDOR is not None
