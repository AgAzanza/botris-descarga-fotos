"""
comun.py — lo que comparten todos los programas de marca.

Aca vive lo que NO cambia entre marcas: leer el excel, armar el codigo del
producto, comparar fotos para detectar repetidas, bajar, y generar la hoja
de revision.

Lo que SI cambia entre marcas (de donde salen las fotos, como se arma la
URL, en que orden van) vive en el archivo de cada marca: geox.py,
lanidor.py, etc.

Ningun archivo de marca deberia necesitar copiar codigo de otro. Si algo
sirve para dos marcas, va aca.
"""

import io
import re
import sys
import time
import unicodedata
from pathlib import Path

# En Windows la consola usa cp1252 y revienta al imprimir un nombre de
# producto con acentos. Esto la pasa a UTF-8; en Linux y Mac no hace nada.
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import pandas as pd
import requests
from PIL import Image, ImageChops, ImageOps


# ===========================================================================
# LECTURA DEL EXCEL
# ===========================================================================

# Los excels de Comex traen la misma info con nombres de columna distintos
# segun la marca. Se prueban en orden; el primero que exista, gana.
COLUMNAS = {
    "codigo":     ["CODIGO"],
    "nombre":     ["NOMBRE"],
    "marca":      ["COD.MARCA", "MARCA"],
    "referencia": ["REFERENCIA", "MATERIAL"],   # Hugo la llama MATERIAL
    "cod_color":  ["COD.COLOR"],
    "color":      ["COLOR", "COLOR LETRAS"],
    "talla":      ["COD.TALLA", "TALLA", "TALLA HB"],
    "ean":        ["COD.BARRA"],
    "coleccion":  ["COD.COLECCION", "COLECCION", "COLECCIÓN"],
}

OBLIGATORIAS = ["codigo", "nombre", "marca", "talla"]


def leer(path):
    """
    Lee la primera hoja del excel, entera como texto.

    dtype=str porque si pandas infiere tipos, COD.COLOR "001" se vuelve 1 y
    el EAN 8058192734342 se vuelve 8.05e+12. Ninguno sirve despues.

    sheet_name=0 porque cada archivo llama a su hoja distinto (1748, 1448...).
    """
    path = Path(path)
    motor = "xlrd" if path.suffix.lower() == ".xls" else "openpyxl"
    df = pd.read_excel(path, sheet_name=0, engine=motor,
                       dtype=str, keep_default_na=False)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def unificar_columnas(df):
    """Renombra las columnas al nombre canonico y limpia espacios."""
    disponibles = {c.upper(): c for c in df.columns}
    out = pd.DataFrame(index=df.index)
    for canonico, alias in COLUMNAS.items():
        for a in alias:
            real = disponibles.get(a.upper())
            if real is not None:
                out[canonico] = df[real].astype(str).str.strip()
                break

    faltan = [c for c in OBLIGATORIAS if c not in out.columns]
    if faltan:
        raise ValueError(f"Faltan columnas: {faltan}")

    # MK trae nombres con espacios de relleno ("MK POOL SLIDE      ")
    # y Lanidor con dobles espacios.
    out["nombre"] = out["nombre"].str.replace(r"\s+", " ", regex=True).str.strip()
    return out[out["codigo"] != ""]


# ===========================================================================
# CODIGO DEL PRODUCTO  (= nombre del archivo)
# ===========================================================================

def limpio(texto):
    """Deja solo letras y numeros, en mayusculas. '400549.118' -> '400549118'"""
    t = unicodedata.normalize("NFKD", str(texto))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"[^A-Za-z0-9]", "", t).upper()


def codigo_producto(codigo, talla):
    """
    El CODIGO del excel sin la talla y sin puntos. Es el nombre del archivo
    y tambien la clave con la que se agrupan las filas.

        D650ZB00022C400535   talla 35    -> D650ZB00022C4005
        D650ZB00022C4005365  talla 36.5  -> D650ZB00022C4005
        400549.118.L         talla L     -> 400549118
        MT61020LC5657OS      talla O/S   -> MT61020LC5657
        50507803O271L        talla L     -> 50507803O271

    OJO CON EL ORDEN: primero se limpian los puntos y barras, DESPUES se
    recorta la talla. Al reves falla, porque Geox escribe la talla 36.5 en
    la columna pero 365 dentro del codigo, y MK escribe O/S pero OS.
    """
    c, t = limpio(codigo), limpio(talla)
    return c[:-len(t)] if t and c.endswith(t) and len(c) > len(t) else c


def productos(df):
    """
    Colapsa las filas-talla en productos, agrupando por codigo_producto.

    Se agrupa por el codigo y NO por referencia+color porque en el excel de
    MK los cinturones traen un EAN distinto por talla en la columna
    REFERENCIA: agrupando por referencia salian 4 productos donde hay 1.
    Verificado en los 5 excels: dentro de cada codigo, nombre y color son
    siempre uno solo.
    """
    df = df.copy()
    df["cod"] = [codigo_producto(c, t) for c, t in zip(df["codigo"], df["talla"])]

    agg = {c: "first" for c in
           ("nombre", "color", "marca", "referencia", "cod_color", "ean", "coleccion")
           if c in df.columns}

    prod = df.groupby("cod", as_index=False, sort=False).agg(**{
        k: pd.NamedAgg(column=k, aggfunc=v) for k, v in agg.items()
    })

    tallas = (df.groupby("cod", sort=False)["talla"]
                .apply(lambda s: "/".join(dict.fromkeys(x for x in s if x)))
                .reset_index(name="tallas"))
    return prod.merge(tallas, on="cod", how="left")


def cargar(path, marca=None):
    """Excel -> tabla de productos. Si se pasa marca, filtra por esa marca."""
    prod = productos(unificar_columnas(leer(path)))
    if marca:
        prod = prod[prod["marca"].str.upper() == marca.upper()]
    return prod.reset_index(drop=True)


def _columna_real(df, canonico):
    """El nombre que tiene en ESTE excel una columna canonica."""
    disponibles = {c.upper(): c for c in df.columns}
    for alias in COLUMNAS.get(canonico, []):
        real = disponibles.get(alias.upper())
        if real is not None:
            return real
    return None


def excel_encontrados(origen, rep, carpeta_marca, marca=None,
                      nombre="encontrados.xlsx"):
    """
    Escribe un excel con las filas de los productos que SI bajaron fotos.

    Se relee el archivo original y se filtran sus filas, en vez de armar uno
    nuevo a partir de los datos que maneja el programa. De esa forma salen
    TODAS las columnas tal cual venian, con sus nombres y su orden, y el
    cliente puede usarlo como usaria el original.

    Van todas las filas-talla de cada producto encontrado, no una por
    producto: el que lo reciba espera la misma estructura que mando.

    Devuelve (ruta, filas, productos) o (None, 0, 0) si no hubo ninguno.
    """
    if rep is None or len(rep) == 0 or "cod" not in rep.columns:
        return None, 0, 0

    n = pd.to_numeric(rep.get("n_fotos", 0), errors="coerce").fillna(0)
    codigos = set(rep.loc[n > 0, "cod"].astype(str))
    if not codigos:
        return None, 0, 0

    crudo = leer(origen)
    col_codigo = _columna_real(crudo, "codigo")
    col_talla = _columna_real(crudo, "talla")
    if col_codigo is None or col_talla is None:
        return None, 0, 0

    cods = [codigo_producto(c, t)
            for c, t in zip(crudo[col_codigo], crudo[col_talla])]
    quedan = pd.Series(cods, index=crudo.index).isin(codigos)

    # Si se paso la marca, solo sus filas: cada marca tiene su propio excel.
    col_marca = _columna_real(crudo, "marca")
    if marca and col_marca is not None:
        quedan &= crudo[col_marca].astype(str).str.strip().str.upper() == marca.upper()

    sub = crudo[quedan]
    if len(sub) == 0:
        return None, 0, 0

    destino = Path(carpeta_marca) / nombre
    destino.parent.mkdir(parents=True, exist_ok=True)
    sub.to_excel(destino, index=False)
    return destino, len(sub), len(codigos)


# ===========================================================================
# COMPARACION DE FOTOS  (detectar repetidas)
# ===========================================================================

LADO_FIRMA = 96        # a cuantos pixeles se reduce cada foto para comparar
UMBRAL_FONDO = 238     # por encima de este gris se considera fondo y se recorta
DIF_IGUAL = 5          # diferencia media por pixel bajo la cual son la misma


def firma(img, lado=LADO_FIRMA):
    """
    Reduce la foto a una imagen chica comparable, en tres movimientos:

    1. RECORTA EL FONDO. En una foto de catalogo el producto ocupa poco y el
       resto es blanco. Sin recortar, dos angulos distintos del mismo zapato
       quedan los dos como "mancha oscura al medio sobre blanco" y se
       confunden.
    2. NORMALIZA EL TAMANO, para que la misma foto en 1024 y en 2048 de igual.
    3. AUTOCONTRASTE, para absorber diferencias de compresion.
    """
    g = img.convert("L")
    mascara = g.point(lambda p: 0 if p > UMBRAL_FONDO else 255)
    caja = mascara.getbbox()
    if caja and (caja[2] - caja[0]) > 20 and (caja[3] - caja[1]) > 20:
        g = g.crop(caja)
    return ImageOps.autocontrast(g.resize((lado, lado), Image.LANCZOS))


def diferencia(fa, fb):
    """
    Diferencia media por pixel entre dos firmas. 0 = identicas.

    Medido sobre fotos reales del D XAND 3 avio:
        la misma foto en 2000 y en 1024 webp   ->   2.3   (se fusionan)
        dos angulos distintos del mismo zapato ->  10.8   (se guardan las 2)
    """
    d = ImageChops.difference(fa, fb).tobytes()
    return sum(d) / len(d)


def unicas(imagenes):
    """
    Recibe [(orden, imagen), ...] y descarta las repetidas, quedandose con
    la de mayor resolucion de cada grupo. Devuelve las imagenes ordenadas.
    """
    salida = []
    for orden, img in imagenes:
        f = firma(img)
        px = img.width * img.height
        gemela = next((u for u in salida
                       if diferencia(u["firma"], f) <= DIF_IGUAL), None)
        if gemela is None:
            salida.append({"orden": orden, "img": img, "firma": f, "px": px})
        elif px > gemela["px"]:
            gemela.update(img=img, firma=f, px=px)
    salida.sort(key=lambda u: u["orden"])
    return [u["img"] for u in salida]


# ===========================================================================
# RED
# ===========================================================================

HEADERS = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64)"}
MINIMO_BYTES = 5000
ESPERA = 0.15


def existe(url, sesion):
    """HEAD: pregunta si la foto esta, sin bajarla."""
    try:
        r = sesion.head(url, headers=HEADERS, timeout=20, allow_redirects=True)
        return r.status_code == 200
    except requests.RequestException:
        return False
    finally:
        time.sleep(ESPERA)


def bajar(url, sesion):
    """Devuelve la imagen PIL, o None si no esta o es un placeholder."""
    try:
        r = sesion.get(url, headers=HEADERS, timeout=30)
    except requests.RequestException:
        return None
    finally:
        time.sleep(ESPERA)
    if r.status_code != 200 or len(r.content) < MINIMO_BYTES:
        return None
    try:
        img = Image.open(io.BytesIO(r.content))
        img.load()
        return img
    except Exception:
        return None


# Dentro de la carpeta de cada marca, las imagenes van en su propia
# subcarpeta. Asi lo que se le entrega al cliente es una sola carpeta con
# fotos y nada mas, y el material de control queda aparte:
#
#     fotos/Geox/fotos/      <- esto se entrega
#     fotos/Geox/revision/   <- esto es interno
#
# Si se quiere otro nombre, se cambia aca y listo.
CARPETA_FOTOS = "fotos"


def carpeta_fotos(carpeta_marca):
    return Path(carpeta_marca) / CARPETA_FOTOS


def nombre_foto(codigo, indice=0, sufijo="", ext="webp"):
    """
    El nombre de archivo de una foto.

        indice 0 ->  D650ZB00022C4005-geox-ecuador.webp
        indice 1 ->  D650ZB00022C4005-1-geox-ecuador.webp

    El sufijo de marca va AL FINAL y el numero de foto en el medio. No es
    un capricho: es el formato que espera el archivo de carga de
    WooCommerce, asi que el archivo que se sube y la URL que va al CSV son
    exactamente el mismo texto, sin renombrar nada en el medio.

    >>> nombre_foto("405608188", 0, "lanidor-ecuador")
    '405608188-lanidor-ecuador.webp'
    >>> nombre_foto("405608188", 2, "lanidor-ecuador")
    '405608188-2-lanidor-ecuador.webp'
    >>> nombre_foto("405608188", 1)
    '405608188-1.webp'
    """
    partes = [str(codigo)]
    if indice:
        partes.append(str(indice))
    if sufijo:
        partes.append(str(sufijo))
    return "-".join(partes) + "." + ext


def guardar(imagenes, carpeta_marca, codigo, sufijo="", calidad=90):
    """
    Guarda las fotos de un producto en la subcarpeta de imagenes.

        fotos/Geox/fotos/D650ZB00022C4005-geox-ecuador.webp      <- principal
        fotos/Geox/fotos/D650ZB00022C4005-1-geox-ecuador.webp
        fotos/Geox/fotos/D650ZB00022C4005-2-geox-ecuador.webp
    """
    if not imagenes:
        return []
    cp = carpeta_fotos(carpeta_marca)
    cp.mkdir(parents=True, exist_ok=True)
    nombres = []
    for n, img in enumerate(imagenes):
        destino = cp / nombre_foto(codigo, n, sufijo)
        img.convert("RGB").save(destino, "WEBP", quality=calidad)
        nombres.append(destino.name)
    return nombres


def ya_bajado(carpeta_marca, codigo, sufijo=""):
    return (carpeta_fotos(carpeta_marca) / nombre_foto(codigo, 0, sufijo)).exists()


def archivos_de(carpeta_marca, codigo, sufijo=""):
    """
    Las fotos ya bajadas de un producto, en orden: la principal primero.

    Se listan explicitamente y no con glob(f"{codigo}*") para no capturar
    las de otro producto cuyo codigo empiece igual.
    """
    cp = carpeta_fotos(carpeta_marca)

    # Si no estan con el sufijo, se buscan sin el: las fotos bajadas antes de
    # que existiera el sufijo se llaman {codigo}.webp a secas y no hay motivo
    # para obligar a rebajarlas.
    usar = sufijo
    if sufijo and not (cp / nombre_foto(codigo, 0, sufijo)).exists() \
            and (cp / nombre_foto(codigo, 0, "")).exists():
        usar = ""

    salida = []
    if (cp / nombre_foto(codigo, 0, usar)).exists():
        salida.append(nombre_foto(codigo, 0, usar))
    n = 1
    while (cp / nombre_foto(codigo, n, usar)).exists():
        salida.append(nombre_foto(codigo, n, usar))
        n += 1
    return salida


def renombrar_con_sufijo(carpeta_marca, sufijo):
    """
    Le agrega el sufijo de marca a las fotos que no lo tengan.

    Para las fotos bajadas antes de que el sufijo existiera: evita tener que
    volver a descargarlas solo por el nombre.

    Devuelve (renombradas, ya_estaban).
    """
    cp = carpeta_fotos(carpeta_marca)
    if not cp.exists() or not sufijo:
        return 0, 0

    hechas = saltadas = 0
    for f in sorted(cp.glob("*.webp")):
        if f.stem.endswith(sufijo):
            saltadas += 1
            continue
        destino = cp / f"{f.stem}-{sufijo}.webp"
        if destino.exists():
            saltadas += 1
            continue
        f.rename(destino)
        hechas += 1
    return hechas, saltadas


# ===========================================================================
# HOJA DE REVISION
# ===========================================================================

_CSS = """
:root { --fondo:#9aa0a6; --panel:#eceef0; --tinta:#1c1f22; --suave:#5d646b; --alerta:#a3341f; }
* { box-sizing:border-box; }
body { margin:0; padding:24px; background:var(--fondo); color:var(--tinta);
       font:15px/1.5 ui-sans-serif, system-ui, -apple-system, sans-serif; }
header { max-width:1400px; margin:0 auto 20px; }
h1 { font-size:21px; font-weight:600; margin:0 0 4px; }
.resumen { color:#2b3035; margin:0 0 14px; max-width:70ch; }
.filtros { display:flex; gap:8px; flex-wrap:wrap; }
.filtros button { font:inherit; font-size:14px; padding:6px 14px; cursor:pointer;
  border:1px solid #6f767d; border-radius:3px; background:var(--panel); color:var(--tinta); }
.filtros button[aria-pressed="true"] { background:var(--tinta); color:#fff; border-color:var(--tinta); }
main { max-width:1400px; margin:0 auto; }
.producto { display:grid; grid-template-columns:270px 1fr; gap:20px;
  background:var(--panel); border-radius:4px; padding:16px; margin-bottom:12px; }
.producto.sinfotos { border-left:5px solid var(--alerta); }
.ficha .codigo { font-family:ui-monospace,"SF Mono",Menlo,monospace;
  font-size:15px; font-weight:600; }
.ficha .nombre { margin-top:4px; }
.ficha .color { color:var(--suave); }
.ficha .cuenta { margin-top:8px; font-size:13px; }
.aviso { color:var(--alerta); font-weight:600; }
.nota { margin-top:6px; font-size:13px; color:var(--suave); }
.producto.sinfotos .nota { color:var(--alerta); }
.links { margin-top:10px; display:flex; flex-direction:column; gap:4px; }
.links a { color:#1f4e79; font-size:13px; }
.tira { display:flex; gap:10px; flex-wrap:wrap; }
.foto img { width:150px; height:150px; object-fit:contain; background:#fff;
  border:1px solid #c8ccd0; border-radius:3px; display:block; }
.foto span { display:block; font-size:12px; color:var(--suave); text-align:center; margin-top:3px; }
.vacio { color:var(--suave); font-style:italic; align-self:center; }
@media (max-width:720px) { .producto { grid-template-columns:1fr; } }
"""

_JS = """
const bs = document.querySelectorAll('.filtros button');
bs.forEach(b => b.addEventListener('click', () => {
  bs.forEach(o => o.setAttribute('aria-pressed', o === b));
  const f = b.dataset.filtro;
  document.querySelectorAll('.producto').forEach(p => {
    const n = Number(p.dataset.fotos);
    p.style.display = (f==='todos' || (f==='sin'&&n===0) || (f==='pocas'&&n>0&&n<=2)) ? '' : 'none';
  });
}));
"""


def hoja_revision(rep, carpeta_marca, marca,
                  titulo_links=("Ver en el sitio", "Foto original")):
    """
    Genera la hoja de revision de UNA marca, en su propia subcarpeta:

        fotos/Geox/fotos/D650ZB00022C4005.webp   <- las fotos
        fotos/Geox/revision/revision.html        <- la hoja
        fotos/Geox/revision/reporte.csv

    Va en subcarpeta por dos motivos: cada marca tiene la suya y no se
    pisan entre si, y la carpeta que se le entrega al cliente queda con
    fotos solamente.

    Como la hoja esta un nivel mas abajo, las imagenes se referencian
    con ../

    rep: DataFrame con columnas cod, nombre, color, n_fotos, archivos,
         url_producto, url_foto1.
    """
    import html as _h

    rep = rep.copy()
    rep["n_fotos"] = pd.to_numeric(rep["n_fotos"], errors="coerce").fillna(0).astype(int)
    rep = rep.sort_values("n_fotos", kind="stable")   # los problemas arriba

    bloques, total = [], 0
    for _, f in rep.iterrows():
        archivos = str(f.get("archivos", "")).split()
        total += len(archivos)

        tira = "".join(
            f'<div class="foto"><a href="../{CARPETA_FOTOS}/{_h.escape(a)}" target="_blank">'
            f'<img src="../{CARPETA_FOTOS}/{_h.escape(a)}" loading="lazy" alt=""></a>'
            f'<span>{i}</span></div>'
            for i, a in enumerate(archivos, 1)
        ) or '<p class="vacio">Sin fotos. Puede que el producto no este publicado en el sitio.</p>'

        links = []
        if f.get("url_producto"):
            # cada fila puede traer su propia etiqueta: la ficha exacta cuando
            # se encontro, la busqueda cuando no.
            etiqueta = str(f.get("label_producto") or titulo_links[0])
            links.append(f'<a href="{_h.escape(str(f["url_producto"]))}" target="_blank" '
                         f'rel="noopener">{_h.escape(etiqueta)}</a>')
        if f.get("url_foto1"):
            links.append(f'<a href="{_h.escape(str(f["url_foto1"]))}" target="_blank" '
                         f'rel="noopener">{titulo_links[1]}</a>')

        n = len(archivos)
        cuenta = f"<b>{n}</b> fotos" if n else '<span class="aviso">sin fotos</span>'
        nota = (f'<div class="nota">{_h.escape(str(f["nota"]))}</div>'
                if f.get("nota") else "")
        bloques.append(f"""
    <section class="producto{' sinfotos' if not n else ''}" data-fotos="{n}">
      <div class="ficha">
        <div class="codigo">{_h.escape(str(f['cod']))}</div>
        <div class="nombre">{_h.escape(str(f.get('nombre','')))}</div>
        <div class="color">{_h.escape(str(f.get('color','')))}</div>
        <div class="cuenta">{cuenta}</div>
        {nota}
        <div class="links">{" ".join(links)}</div>
      </div>
      <div class="tira">{tira}</div>
    </section>""")

    sin = int((rep.n_fotos == 0).sum())
    pocas = int(((rep.n_fotos > 0) & (rep.n_fotos <= 2)).sum())

    doc = f"""<!doctype html>
<html lang="es"><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Revision de fotos — {marca}</title>
<style>{_CSS}</style>
<header>
  <h1>Revision de fotos — {marca}</h1>
  <p class="resumen">{len(rep)} productos, {total} fotos. Revisa que el color de
  la foto coincida con el de la ficha, que no haya repetidas y que ninguna sea
  de otro producto. Los que no tienen fotos llevan un link al buscador del
  sitio, para poder comprobar si el producto no existe o si le falta ese
  color.</p>
  <div class="filtros">
    <button data-filtro="todos" aria-pressed="true">Todos ({len(rep)})</button>
    <button data-filtro="sin" aria-pressed="false">Sin fotos ({sin})</button>
    <button data-filtro="pocas" aria-pressed="false">1 o 2 fotos ({pocas})</button>
  </div>
</header>
<main>{"".join(bloques)}</main>
<script>{_JS}</script></html>"""

    rev = Path(carpeta_marca) / "revision"
    rev.mkdir(parents=True, exist_ok=True)
    destino = rev / "revision.html"
    destino.write_text(doc, encoding="utf-8")
    # utf-8-sig y no utf-8 a secas: sin el BOM, Excel en Windows abre el
    # CSV con los acentos rotos.
    rep.to_csv(rev / "reporte.csv", index=False, encoding="utf-8-sig")
    return destino, total, sin