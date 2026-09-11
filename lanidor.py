#!/usr/bin/env python3
"""
LANIDOR — baja las fotos de los productos de un excel de estructura.

Lo especifico de esta marca vive aca; lo compartido esta en comun.py.

Como se consiguen las fotos
---------------------------
Lanidor no permite construir la URL de la foto desde el excel: los codigos
de la web no tienen nada que ver con los del excel. Pero el sitio tiene un
endpoint JSON que resuelve todo de una:

    POST /listaprodutos.aspx/getlistaProdutos
    {"search": "405594", ...}

Devuelve un item POR COLOR, y cada uno trae su propia lista de fotos:

    "Referencia": 405594,  "NumeroCor": 383,
    "arrayVistas": "184302,184303,184304,184305,184306,",
    "versaoCDN": 1,  "extensaoImgs": "jpg"

Eso resuelve el problema que teniamos: separar las fotos por color. El
endpoint ya las entrega separadas.

La URL de cada foto se arma segun versaoCDN (formula sacada del JS del
propio sitio):

    versaoCDN = 1 -> lanidor.com/cdn-srcs/products/{vista}.{ext}
    versaoCDN = 0 -> cdn.lanidor.com/Produtos/grandes/{idProduto}_{vista}_{g|m}.{ext}
                     (la letra depende de si el producto es anterior o
                      posterior al 25/02/2019)

UNA LLAMADA POR REFERENCIA
--------------------------
La busqueda devuelve TODOS los colores de la referencia, asi que se pide
una vez por referencia y no una vez por producto. En el excel de Lanidor
eso son 153 llamadas en vez de 195.

OJO: el sitio responde 403 a curl pelado. Hay que mandar User-Agent.

USO
---
    pip install pandas xlrd requests pillow

    python lanidor.py Lanidor.xls --limite 3
    python lanidor.py Lanidor.xls
    python lanidor.py Lanidor.xls --marca FUSTER
"""

import argparse
import json
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd
import requests

import comun


# ===========================================================================
# SESION  (Cloudflare)
# ===========================================================================
#
# lanidor.com esta detras de Cloudflare y responde con una pagina de
# desafio ("Just a moment...") a requests. No es cuestion de cabeceras:
# Cloudflare mira la huella TLS de la libreria y `requests` se delata.
#
# curl_cffi imita la huella de Chrome y normalmente pasa:
#
#     pip install curl_cffi
#
# Si aun asi no pasa, se le pueden dar las cookies del navegador con
# --cookies (ver la ayuda del parametro).

try:
    from curl_cffi import requests as cffi
    HAY_CFFI = True
except ImportError:
    HAY_CFFI = False


def nueva_sesion(cookies=None):
    if HAY_CFFI:
        s = cffi.Session(impersonate="chrome")
    else:
        s = requests.Session()
        print("  aviso: curl_cffi no esta instalado. Con requests a secas,\n"
              "         Cloudflare va a devolver la pagina de desafio.\n"
              "         pip install curl_cffi\n")
    if cookies:
        for parte in cookies.split(";"):
            if "=" in parte:
                k, v = parte.split("=", 1)
                s.cookies.set(k.strip(), v.strip(), domain=".lanidor.com")
    return s


# ===========================================================================
# LO ESPECIFICO DE LANIDOR
# ===========================================================================

MARCA = "LANI"
CARPETA = "Lanidor"

BUSQUEDA = "https://www.lanidor.com/listaprodutos.aspx/getlistaProdutos"
FICHA = "https://www.lanidor.com{path}"

# Para los productos que no se encontraron: el buscador publico del sitio.
# Es el link que le permite al cliente comprobar por si mismo si el producto
# no existe o si simplemente no tiene ese color.
BUSCADOR_WEB = "https://www.lanidor.com/uk/search/{referencia}"

# El endpoint espera TODOS estos campos, aunque casi todos vayan vacios.
# Copiado del JS de la propia pagina (funcion callProductsList).
CAMPOS_VACIOS = [
    "idDestaque", "op", "opcol", "idTipoCategoria", "idTipoMenu",
    "idTipoMenuNot", "ipag", "srchop", "srchtpa", "srchcat", "srchestilo",
    "idTamanho", "tgcor", "tpmnu", "zona", "z", "tpa", "brand", "newcol",
    "estacao", "brands", "pc", "idTipoEstilo", "mnu", "filtroTipoCaract",
    "filtroCaract", "filtroTamanhos", "filtroCores", "filtroOrdem", "refs",
    "idsubtipomenu",
]

CABECERAS = {
    "Content-Type": "application/json; charset=utf-8",
    "User-Agent": ("Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/120.0 Safari/537.36"),
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "X-Requested-With": "XMLHttpRequest",
    "Referer": "https://www.lanidor.com/",
}

# Fecha a partir de la cual las fotos viejas usan sufijo "m" en vez de "g".
CORTE_M = datetime(2019, 2, 25, 11, tzinfo=timezone.utc)


class Bloqueado(RuntimeError):
    """Cloudflare respondio con el desafio en vez de los datos."""


def buscar_referencia(referencia, sesion):
    """
    Pide al sitio todos los colores de una referencia.

    Devuelve la lista de items tal cual los manda el servidor, ya filtrada
    para quedarse solo con los de ESTA referencia (la busqueda puede traer
    otros productos parecidos).
    """
    cuerpo = {"search": str(referencia), "numpag": 1,
              "sitemarca": "1", "sitebeta": "0", "site2bstyle": "0"}
    cuerpo.update({c: "" for c in CAMPOS_VACIOS})

    try:
        r = sesion.post(BUSQUEDA, headers=CABECERAS,
                        data=json.dumps(cuerpo), timeout=30)
    except requests.RequestException:
        return []
    if r.status_code != 200:
        return []

    if "Just a moment" in r.text[:400] or "cf-browser-verification" in r.text[:2000]:
        raise Bloqueado(
            "Cloudflare devolvio la pagina de desafio en vez de los datos.\n"
            "  - si no lo tenes:  pip install curl_cffi\n"
            "  - si ya lo tenes:  pasa las cookies del navegador con --cookies"
        )

    try:
        # La respuesta es {"d": "<json como texto>"}: hay que parsear 2 veces.
        items = json.loads(r.json()["d"])
    except Exception:
        return []

    objetivo = comun.limpio(referencia)
    return [it for it in items
            if comun.limpio(it.get("Referencia", "")) == objetivo]


def urls_de_item(item):
    """Arma las URLs de las fotos de un color, segun la version de CDN."""
    ext = item.get("extensaoImgs") or "jpg"
    vistas = [v.strip() for v in str(item.get("arrayVistas") or "").split(",")
              if v.strip()]
    if not vistas:
        return []

    if item.get("versaoCDN") == 1:
        return [f"https://www.lanidor.com/cdn-srcs/products/{v}.{ext}"
                for v in vistas]

    idp = item.get("idProduto")
    letra = "g"
    fecha = str(item.get("DataInsercao") or "")
    if "Date(" in fecha:
        try:
            ms = int(fecha.split("Date(")[1].split(")")[0])
            if datetime.fromtimestamp(ms / 1000, tz=timezone.utc) >= CORTE_M:
                letra = "m"
        except Exception:
            pass
    return [f"https://cdn.lanidor.com/Produtos/grandes/{idp}_{v}_{letra}.{ext}"
            for v in vistas]


# El navegador carga las fotos a traves del redimensionador de Cloudflare:
#
#   imgs.lanidor.com/cdn-cgi/image/width=800,quality=90,format=webp/{original}
#   \_________ redimensionador _________/\________ la foto real ________/
#
# Nosotros bajamos el ORIGINAL, que viene a resolucion completa en vez de
# recortado a 800 o 500 px. Si algun dia el origen deja de responder, se
# reintenta por el redimensionador pidiendo un ancho grande.

RESIZER = ("https://imgs.lanidor.com/cdn-cgi/image/"
           "width={ancho},quality=95/{original}")


def via_resizer(url, ancho=2000):
    return RESIZER.format(ancho=ancho, original=url)


def indexar(items):
    """
    { codigo_de_color: item }  para poder cruzar con el COD.COLOR del excel.

    El excel trae el color como texto ("383", "047", "8") y el JSON como
    numero (383). Se normalizan los dos a entero-en-texto para que matcheen.
    """
    salida = {}
    for it in items:
        cor = str(it.get("NumeroCor", "")).strip()
        if not cor:
            continue
        claves = {cor, cor.lstrip("0") or "0"}
        try:
            claves.add(str(int(cor)))
        except ValueError:
            pass
        for k in claves:
            salida[k] = it
    return salida


def claves_color(cod_color):
    """Las formas en que un COD.COLOR del excel puede aparecer en el JSON."""
    c = str(cod_color).strip()
    formas = {c, c.lstrip("0") or "0"}
    try:
        formas.add(str(int(c)))
    except ValueError:
        pass
    return formas


# ===========================================================================
# PRINCIPAL
# ===========================================================================

def main(argv=None):
    """argv=None usa la linea de comandos; una lista permite llamarlo
    desde codigo (lo hace runner.py cuando corre dentro del .exe)."""
    ap = argparse.ArgumentParser(description="Baja las fotos de Lanidor.")
    ap.add_argument("excel", type=Path)
    ap.add_argument("--salida", type=Path, default=Path("fotos"))
    ap.add_argument("--limite", type=int, default=None)
    ap.add_argument("--marca", default=MARCA,
                    help="LANI (por defecto) o FUSTER")
    ap.add_argument("--forzar", action="store_true")
    ap.add_argument("--cookies", default=None,
                    help="cookies del navegador si Cloudflare sigue bloqueando. "
                         "En el navegador: F12 > Network > cualquier pedido a "
                         "lanidor.com > Request Headers > copiar el valor de "
                         "Cookie. Hace falta cf_clearance. Caduca en unas horas "
                         "y esta atado a tu IP.")
    args = ap.parse_args(argv)

    prod = comun.cargar(args.excel, marca=args.marca)
    if len(prod) == 0:
        todas = comun.cargar(args.excel)["marca"].unique()
        print(f"El excel no tiene productos {args.marca}. Marcas: {list(todas)}")
        return
    if args.limite:
        prod = prod.head(args.limite)

    carpeta = args.salida / (CARPETA if args.marca == MARCA else args.marca.title())
    carpeta.mkdir(parents=True, exist_ok=True)

    refs = prod["referencia"].nunique()
    print(f"\n{args.marca}: {len(prod)} productos, {refs} referencias -> {carpeta}/")
    print(f"una consulta por referencia, no por producto\n")

    sesion = nueva_sesion(args.cookies)
    cache = {}          # referencia -> { color: item }
    filas = []

    for i, (_, p) in enumerate(prod.iterrows(), 1):
        cod, ref = p["cod"], p["referencia"]

        if not args.forzar and comun.ya_bajado(carpeta, cod):
            existentes = comun.archivos_de(carpeta, cod)
            print(f"  {i:>3}/{len(prod)} [--] {cod:<16} ya estaba ({len(existentes)})")
            filas.append(_fila(p, existentes, 0, None))
            continue

        if ref not in cache:
            try:
                cache[ref] = indexar(buscar_referencia(ref, sesion))
            except Bloqueado as e:
                print(f"\n  BLOQUEADO en la referencia {ref}:\n  {e}\n")
                return
            comun.time.sleep(comun.ESPERA)

        item = next((cache[ref][k] for k in claves_color(p["cod_color"])
                     if k in cache[ref]), None)

        if item is None:
            colores = sorted({str(v.get("NumeroCor")) for v in cache[ref].values()})
            detalle = (f"El sitio tiene la referencia pero no el color "
                       f"{p['cod_color']}. Colores publicados: "
                       f"{', '.join(colores)}." if colores
                       else "La referencia no aparece en el buscador del sitio.")
            print(f"  {i:>3}/{len(prod)} [!!] {cod:<16} {detalle}")
            filas.append(_fila(p, [], 0, None, nota=detalle))
            continue

        crudas = []
        for n, url in enumerate(urls_de_item(item)):
            img = comun.bajar(url, sesion)
            if img is None:
                # el origen no respondio: probamos por el redimensionador
                img = comun.bajar(via_resizer(url), sesion)
            if img is not None:
                crudas.append((n, img))

        imagenes = comun.unicas(crudas)
        nombres = comun.guardar(imagenes, carpeta, cod)
        repetidas = len(crudas) - len(imagenes)

        nota = f"{len(nombres)} fotos" if nombres else "sin fotos"
        if repetidas:
            nota += f"  ({repetidas} repetidas descartadas)"
        print(f"  {i:>3}/{len(prod)} {'[ok]' if nombres else '[!!]'} {cod:<16} {nota}")

        filas.append(_fila(p, nombres, repetidas, item))

    rep = pd.DataFrame(filas)
    destino, total, sin = comun.hoja_revision(rep, carpeta, carpeta.name)
    xls, n_filas, n_prod = comun.excel_encontrados(args.excel, rep, carpeta,
                                                   marca=args.marca)

    print(f"\n{'-' * 64}")
    print(f"  productos con fotos  : {(rep.n_fotos > 0).sum()} de {len(rep)}")
    print(f"  fotos guardadas      : {total}")
    print(f"  repetidas descartadas: {rep.repetidas.sum()}")
    print(f"  consultas al sitio   : {len(cache)}")
    if xls:
        print(f"  excel de encontrados : {xls.name}  "
              f"({n_prod} productos, {n_filas} filas)")
    print(f"\n  Abri:  {destino.resolve()}")


def _fila(p, nombres, repetidas, item, nota=""):
    urls = urls_de_item(item) if item else []

    # Con ficha encontrada, el link va al color exacto. Sin ficha, va al
    # buscador con la referencia, que es lo que deja comprobar por que falto.
    if item and item.get("Urldetalhe"):
        url_producto = FICHA.format(path=item["Urldetalhe"])
        etiqueta = "Ver la ficha de este color"
    else:
        url_producto = BUSCADOR_WEB.format(referencia=p.get("referencia", ""))
        etiqueta = "Buscar la referencia en el sitio"

    return {
        "cod": p["cod"],
        "nombre": p.get("nombre", ""),
        "color": p.get("color", ""),
        "referencia": p.get("referencia", ""),
        "cod_color": p.get("cod_color", ""),
        "tallas": p.get("tallas", ""),
        "n_fotos": len(nombres),
        "repetidas": repetidas,
        "archivos": " ".join(nombres),
        "url_producto": url_producto,
        "label_producto": etiqueta,
        "nota": nota,
        "url_foto1": urls[0] if urls else "",
    }


if __name__ == "__main__":
    main()