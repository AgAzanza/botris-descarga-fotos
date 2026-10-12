#!/usr/bin/env python3
"""
HARMONT & BLAINE — baja las fotos de los productos de un excel.


QUE HACE ESTE PROGRAMA, EN CRIOLLO
==================================

Por cada producto del excel abre su ficha en harmontblaine.com, lee de ahi
los nombres de las fotos (en el orden del carrusel) y las baja en alta.

Leer el excel, comparar fotos repetidas y armar la hoja de revision es
igual para todas las marcas y vive en comun.py.


EL CODIGO
---------
Hay dos formatos (ver GRAMATICAS_TEXTO en comun.py):

    ropa:     CRP011013607 400 L       12 de referencia + 3 de color + talla
    zapatos:  EFM252040A60A4 42        todo menos la talla del final

El codigo SIN la talla es el producto, y con eso se arma la ficha:

    https://www.harmontblaine.com/{codigo sin talla}.html
    ej: https://www.harmontblaine.com/BRL001053163854.html


LAS FOTOS
---------
Harmont corre sobre Salesforce Commerce Cloud, como MK y Liu Jo:

    https://www.harmontblaine.com/on/demandware.static/-/Sites-harmont-m-catalog/
        default/dw835b1dce/images/hi-res/CRR043013967M_808_1.jpg
                                         └─ referencia ─┘ └color┘ └n┘

  - EL NOMBRE NO SE PUEDE ARMAR SOLO. Despues de la referencia de 12 suele
    venir una letra o algo mas (M, I, R, C67...) que el excel no trae. Por
    eso los nombres se LEEN de la ficha, y solo si la ficha no da nada se
    prueba adivinando esa letra (ver PROBAR_LETRAS).
  - La marca de version (dw835b1dce) no hace falta.
  - Cada foto se pide primero en hi-res y, si no esta, en large (en el
    producto de ejemplo, la 5 solo estaba en large). Si la foto viene mas
    ancha que MAX_ANCHO, se achica antes de guardarla.


USO
---
    python harmont.py Inventario.xls --limite 10
    python harmont.py Inventario.xls
    python harmont.py Inventario.xls --forzar
    python harmont.py --diagnostico BRL001053163854
"""

import argparse
import re
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pandas as pd
import requests
from requests.adapters import HTTPAdapter

import comun

try:
    from curl_cffi import requests as cffi
    HAY_CFFI = True
except ImportError:
    HAY_CFFI = False


# ===========================================================================
# LO ESPECIFICO DE HARMONT & BLAINE
# ===========================================================================

MARCA = "HARMONT"
CARPETA = "Harmont"
SUFIJO = "harmontblaine-ecuador"

SITIO = "https://www.harmontblaine.com"
FICHA = SITIO + "/{cod}.html"
BUSQUEDA = SITIO + "/search?q={cod}"
CATALOGO_DEFECTO = "Sites-harmont-m-catalog"
FOTO = SITIO + "/on/demandware.static/-/{catalogo}/default/images/{tam}/{nombre}.{ext}"

# Tamanos que entrega el sitio, del mas grande al mas chico.
TAMANOS = ["hi-res", "large"]

MAX_ANCHO = 2000
MAX_FOTOS = 15          # hasta que numero se prueba cuando se adivina
HILOS = 8

# Cuando la ficha no da fotos (producto despublicado, ficha que no carga),
# se prueba la referencia sola y la referencia + cada letra. Son 27 pruebas
# de la foto 1, en paralelo; si alguna responde, se buscan las demas.
PROBAR_LETRAS = [""] + [chr(c) for c in range(ord("A"), ord("Z") + 1)]

# Una foto del catalogo dentro del html de la ficha, con o sin el servicio
# de redimension (/dw/image/v2/...) y con o sin marca de version.
RE_FOTO = re.compile(
    r"demandware\.static/-/(?P<catalogo>Sites-[A-Za-z0-9_-]+-catalog)/[^/\"'\s]+/"
    r"(?:dw[0-9a-f]+/)?images/(?P<tam>hi-res|large|medium|small)/"
    r"(?P<nombre>(?P<base>[A-Z0-9]+)_(?P<color>[A-Z0-9]+)_(?P<n>\d+))"
    r"\.(?P<ext>jpe?g|png|webp)", re.I)


def url_ficha(cod):
    return FICHA.format(cod=cod)


def nueva_sesion():
    """Para la ficha: con la huella de Chrome si se puede (como MK)."""
    if HAY_CFFI:
        return cffi.Session(impersonate="chrome")
    return requests.Session()


def sesion_fotos():
    """Para las fotos: requests comun, que se puede usar en paralelo."""
    s = requests.Session()
    s.mount("https://", HTTPAdapter(pool_connections=HILOS, pool_maxsize=HILOS))
    return s


def es_del_producto(m, ref, color):
    """
    Si una foto encontrada en la ficha es de ESTE producto. La ficha trae
    tambien fotos de otros ("te puede interesar"), asi que se filtra por la
    referencia y, en la ropa, por el color.
    """
    base, col = m.group("base").upper(), m.group("color").upper()
    if color:
        return base.startswith(ref) and col == color
    # zapatos: el excel no separa el color; vale si referencia+color de la
    # foto empieza con el codigo, o al reves
    junto = base + col
    return junto.startswith(ref) or ref.startswith(base)


def _sin_metas(html):
    """
    Quita las etiquetas <meta> y <link>. Ahi va la foto para compartir en
    redes, que no siempre es la portada (con Hugo eso puso la foto
    equivocada como principal).
    """
    return re.sub(r"<(meta|link)\b[^>]*>", " ", html, flags=re.I)


class Descarga:

    def __init__(self, ancho=MAX_ANCHO):
        self.ficha = nueva_sesion()
        self.fotos_s = sesion_fotos()
        self.ancho = ancho
        self.ultimo = {}            # lo que paso con el ultimo producto (diagnostico)

    # ------------------------------------------------------------- ficha
    def leer(self, url):
        """(status, url_final, html). status 0 si no hubo conexion."""
        try:
            r = self.ficha.get(url, headers={"Accept-Language": "it-IT,en;q=0.8"},
                               timeout=30, allow_redirects=True)
            return r.status_code, str(r.url), r.text or ""
        except Exception:
            return 0, "", ""

    def _fotos_en(self, html, ref, color):
        """Las fotos de ESTE producto que nombra una pagina, en orden."""
        vistos, salida = set(), []
        for m in RE_FOTO.finditer(_sin_metas(html)):
            if not es_del_producto(m, ref, color):
                continue
            nombre = m.group("nombre")
            if nombre.upper() in vistos:
                continue
            vistos.add(nombre.upper())
            salida.append((nombre, m.group("catalogo"), m.group("ext")))
        return salida

    def _registrar(self, url, status, final, html, fotos):
        self.ultimo.setdefault("paginas", []).append(
            {"url": url, "status": status, "final": final, "largo": len(html),
             "en_html": len(RE_FOTO.findall(html)), "del_producto": len(fotos),
             "html": html})

    def ficha_real(self, html, ref, color):
        """
        La direccion de la ficha, sacada de los resultados del buscador.

        La ficha de verdad NO es {codigo}.html: lleva un nombre descriptivo
        y el codigo del sitio, que tiene la letra que el excel no trae:
            /uomo/.../camicia-uomo-popeline-...-CRR043013967M808.html
        """
        for m in re.finditer(r'href="([^"]+?([A-Z0-9]+)\.html)"', html):
            pid = m.group(2).upper()
            if pid.startswith(ref) and pid.endswith(color) \
                    and len(pid) <= len(ref) + len(color) + 4:
                url = m.group(1)
                return url if url.startswith("http") else SITIO + url
        return None

    def de_la_ficha(self, cod, ref, color):
        """
        [(nombre, catalogo, ext)] en el orden del carrusel, o [] si no se
        encontro la ficha.

        1. {codigo}.html directo (a veces funciona: BRL001053163854).
        2. El buscador, con el codigo y con la referencia. De los resultados
           se saca la direccion de la ficha real y se lee esa.
        Si nada da la ficha, devuelve al menos las fotos que mostro el
        buscador (suele mostrar solo 2): el resto se completa del catalogo.

        OJO: si el producto no existe, la ficha NO da error, redirige a la
        portada, que tiene fotos de otros productos. Por eso se filtran las
        fotos por referencia y color.
        """
        self.ultimo["paginas"] = []
        url = FICHA.format(cod=cod)
        status, final, html = self.leer(url)
        fotos = self._fotos_en(html, ref, color)
        self._registrar(url, status, final, html, fotos)
        if fotos:
            self.ultimo["ficha_url"] = final
            return fotos

        de_busqueda = []
        for q in dict.fromkeys([cod, ref]):
            url = BUSQUEDA.format(cod=q)
            status, final, html = self.leer(url)
            fotos = self._fotos_en(html, ref, color)
            self._registrar(url, status, final, html, fotos)
            de_busqueda = de_busqueda or fotos
            ficha = self.ficha_real(html, ref, color)
            if ficha:
                status, final, html = self.leer(ficha)
                fotos = self._fotos_en(html, ref, color)
                self._registrar(ficha, status, final, html, fotos)
                if fotos:
                    self.ultimo["ficha_url"] = final
                    return fotos
            if de_busqueda:
                break
        return de_busqueda

    # ------------------------------------------------------------- fotos
    def bajar(self, nombre, catalogo=CATALOGO_DEFECTO, ext="jpg"):
        """La foto en el tamano mas grande que haya, achicada si hace falta."""
        for tam in TAMANOS:
            img = comun.bajar(FOTO.format(catalogo=catalogo, tam=tam,
                                          nombre=nombre, ext=ext), self.fotos_s)
            if img is not None:
                if img.width > self.ancho:
                    img = img.resize((self.ancho, round(img.height * self.ancho / img.width)))
                return img
        return None

    def existe(self, nombre, tam="hi-res", catalogo=CATALOGO_DEFECTO, ext="jpg"):
        """
        Si la foto esta en el catalogo. Con GET y no con HEAD: no todos los
        servidores contestan bien un HEAD, y un falso "no esta" aca haria
        que se pierdan productos sin que nadie se entere.
        """
        url = FOTO.format(catalogo=catalogo, tam=tam, nombre=nombre, ext=ext)
        try:
            r = self.fotos_s.get(url, headers=comun.HEADERS, timeout=20, stream=True)
            ok = (r.status_code == 200
                  and r.headers.get("Content-Type", "").startswith("image"))
            r.close()
            return ok
        except requests.RequestException:
            return False

    def letra(self, ref, color):
        """
        La letra que va despues de la referencia en el nombre de las fotos
        (la M de CRR043013967M_808_1), probando la foto 1 con cada una.
        Funciona aunque el producto ya no este publicado: el catalogo
        conserva las fotos de temporadas pasadas.
        """
        self.ultimo["letras"] = []
        if not color:
            return None

        def prueba(letra):
            nombre = f"{ref}{letra}_{color}_1"
            return letra if any(self.existe(nombre, tam) for tam in TAMANOS) else None

        with ThreadPoolExecutor(HILOS) as ex:
            letras = [l for l in ex.map(prueba, PROBAR_LETRAS) if l is not None]
        self.ultimo["letras"] = letras
        return letras[0] if letras else None

    def lista(self, cod, ref, color):
        """
        Los nombres de todas las fotos del producto, en orden.

        La pagina (ficha o buscador) da el orden del carrusel, pero no
        siempre todas las fotos: el buscador muestra 2 de 5. Por eso, con
        el nombre base que se sepa (por la pagina o probando letras), se
        prueban ademas TODOS los numeros del 1 al MAX_FOTOS en el catalogo.
        Primero van las de la pagina, en su orden; despues el resto, por
        numero.
        """
        self.ultimo = {"origen": "ficha"}
        de_pagina = self.de_la_ficha(cod, ref, color)

        if de_pagina:
            m = re.match(r"(.+)_([A-Z0-9]+)_\d+$", de_pagina[0][0], re.I)
            base, col = m.group(1), m.group(2)
            catalogo = de_pagina[0][1]
        else:
            letra = self.letra(ref, color)
            if letra is None:
                return []
            self.ultimo["origen"] = "catalogo"
            base, col, catalogo = f"{ref}{letra}", color, CATALOGO_DEFECTO

        todas = list(de_pagina)
        ya = {n.upper() for n, _, _ in todas}
        for n in range(1, MAX_FOTOS + 1):
            nombre = f"{base}_{col}_{n}"
            if nombre.upper() not in ya:
                todas.append((nombre, catalogo, "jpg"))
        return todas

    def fotos(self, cod, ref, color):
        """[(orden, imagen)] del producto, en el orden del carrusel."""
        lista = self.lista(cod, ref, color)
        with ThreadPoolExecutor(HILOS) as ex:
            imgs = list(ex.map(lambda t: self.bajar(*t), lista))
        halladas = [(i, img) for i, img in enumerate(imgs) if img is not None]
        self.ultimo["nombres"] = [lista[i][0] for i, _ in halladas]
        return halladas


# ===========================================================================
# DIAGNOSTICO
# ===========================================================================

# Una foto que SI existe (la mandaste de ejemplo). Sirve de control: si
# esta no baja, el problema es la conexion o la forma de armar la direccion;
# si esta baja y las del producto no, el producto no esta en el catalogo.
FOTO_CONTROL = "CRR043013967M_808_1"


def diagnostico(codigo):
    """
    Muestra paso a paso que pasa con UN producto y guarda el html de cada
    pagina consultada en diagnostico-harmont-{codigo}-{n}.html.
    """
    gram = comun.gramatica(MARCA)
    p = gram.partes(codigo)
    if p:
        cod, ref, color = p["cod"], p["referencia"], p["cod_color"]
    else:
        cod = ref = comun.limpio(codigo)
        color = ""
    print(f"\ncodigo de producto: {cod}   referencia: {ref}   color: {color or '-'}")
    print(f"temporada (3a letra de la referencia): {ref[2:3]}")
    print(f"curl_cffi: {'si' if HAY_CFFI else 'NO (instalalo: pip install curl_cffi)'}")

    d = Descarga()
    ctrl = {t: d.existe(FOTO_CONTROL, t) for t in TAMANOS}
    print(f"\nfoto de control {FOTO_CONTROL}: "
          + ", ".join(f"{t} {'OK' if ok else 'NO'}" for t, ok in ctrl.items()))

    lista = d.lista(cod, ref, color)
    for i, pag in enumerate(d.ultimo.get("paginas", []), 1):
        portada = pag["final"].rstrip("/") == SITIO
        print(f"\npagina {i}: {pag['url']}")
        print(f"  respuesta {pag['status']}, {pag['largo']} caracteres")
        print(f"  termino en: {pag['final']}"
              + ("   <- PORTADA: el sitio no tiene esta ficha" if portada else ""))
        print(f"  fotos en la pagina: {pag['en_html']}, de este producto: {pag['del_producto']}")
        destino = Path(f"diagnostico-harmont-{cod}-{i}.html")
        destino.write_text(pag["html"], encoding="utf-8")
        print(f"  html guardado en: {destino.resolve()}")

    print(f"\nficha encontrada: {d.ultimo.get('ficha_url') or 'ninguna'}")
    if d.ultimo.get("origen") == "catalogo":
        print(f"sin ficha: letra hallada probando el catalogo: {d.ultimo.get('letras')}")
    elif not lista:
        print("sin ficha y " + ("sin color en el codigo (zapato): no se puede probar "
                                "el catalogo" if not color else
                                f"ninguna letra responde en el catalogo"))

    print("\ndescarga (la pagina primero, despues los numeros que falten):")
    if not lista:
        print("    nada para bajar")
    for nombre, cat, ext in lista:
        img = d.bajar(nombre, cat, ext)
        if img:
            print(f"    {nombre:<30} {img.size[0]}x{img.size[1]}")
    if lista:
        print("    (los numeros que no aparecen no existen)")


# ===========================================================================
# PRINCIPAL
# ===========================================================================

def main(argv=None):
    """argv=None usa la linea de comandos; una lista permite llamarlo
    desde codigo (lo hace runner.py cuando corre dentro del .exe)."""
    ap = argparse.ArgumentParser(description="Baja las fotos de Harmont & Blaine.")
    ap.add_argument("excel", type=Path, nargs="?")
    ap.add_argument("--salida", type=Path, default=Path("fotos"))
    ap.add_argument("--limite", type=int, default=None)
    ap.add_argument("--ancho", type=int, default=MAX_ANCHO,
                    help="ancho maximo de las fotos guardadas (por defecto 2000)")
    ap.add_argument("--forzar", action="store_true")
    ap.add_argument("--diagnostico", metavar="CODIGO",
                    help="revisa un solo producto y guarda el html de su ficha")
    args = ap.parse_args(argv)

    if args.diagnostico:
        diagnostico(args.diagnostico)
        return
    if args.excel is None:
        ap.error("falta el excel")

    gram = comun.gramatica(MARCA)
    prod = comun.cargar(args.excel, marca=MARCA, gram=gram, controlar=True)
    if len(prod) == 0:
        todas = comun.cargar(args.excel)["marca"].unique()
        print(f"El excel no tiene productos {MARCA}. Marcas: {list(todas)}")
        return
    if args.limite:
        prod = prod.head(args.limite)

    carpeta = args.salida / CARPETA
    carpeta.mkdir(parents=True, exist_ok=True)
    print(f"\n{MARCA}: {len(prod)} productos -> {carpeta}/")
    print("los nombres de las fotos se leen de la ficha de cada producto\n")
    if not HAY_CFFI:
        print("  aviso: curl_cffi no esta instalado; si la ficha no carga, "
              "instalalo con  pip install curl_cffi\n")

    descarga = Descarga(args.ancho)
    filas, adivinados = [], 0

    for i, (_, p) in enumerate(prod.iterrows(), 1):
        cod = p["cod"]

        if not args.forzar and comun.ya_bajado(carpeta, cod, SUFIJO):
            existentes = comun.archivos_de(carpeta, cod, SUFIJO)
            print(f"  {i:>4}/{len(prod)} [--] {cod:<18} ya estaba ({len(existentes)})")
            filas.append(_fila(p, existentes, 0))
            continue

        nota = ""
        if not p.get("referencia"):
            crudas = []
            nota = "El codigo no tiene el formato de Harmont & Blaine."
        else:
            crudas = descarga.fotos(cod, p["referencia"], p.get("cod_color", ""))
            if descarga.ultimo.get("origen") == "catalogo" and crudas:
                adivinados += 1
                nota = ("No esta publicado en la web; las fotos salieron del "
                        "catalogo de la marca.")
            if not crudas:
                nota = ("No aparece en harmontblaine.com ni en su catalogo de "
                        "fotos." + ("" if p.get("cod_color") else
                                    " En los zapatos solo se puede buscar si estan "
                                    "publicados."))

        imagenes = comun.unicas(crudas)
        nombres = comun.guardar(imagenes, carpeta, cod, SUFIJO)
        repetidas = len(crudas) - len(imagenes)

        detalle = f"{len(nombres)} fotos" if nombres else (nota or "sin fotos")
        if repetidas:
            detalle += f"  ({repetidas} repetidas descartadas)"
        print(f"  {i:>4}/{len(prod)} {'[ok]' if nombres else '[!!]'} {cod:<18} {detalle[:70]}")

        filas.append(_fila(p, nombres, repetidas, nota,
                           descarga.ultimo.get("ficha_url", "")))

    rep = pd.DataFrame(filas)
    destino, total, sin = comun.hoja_revision(rep, carpeta, CARPETA)
    xls, n_filas, n_prod = comun.excel_encontrados(args.excel, rep, carpeta,
                                                   marca=MARCA, gram=gram)

    print(f"\n{'-' * 64}")
    print(f"  productos con fotos  : {(rep.n_fotos > 0).sum()} de {len(rep)}")
    print(f"  fotos guardadas      : {total}")
    print(f"  repetidas descartadas: {rep.repetidas.sum()}")
    if adivinados:
        print(f"  no publicados, fotos del catalogo: {adivinados}")
    if xls:
        print(f"  excel de encontrados : {xls.name}  "
              f"({n_prod} productos, {n_filas} filas)")
    print(f"\n  Abre:  {destino.resolve()}")


def _fila(p, nombres, repetidas, nota="", ficha=""):
    cod = p["cod"]
    return {
        "cod": cod,
        "nombre": p.get("nombre", ""),
        "color": p.get("color", ""),
        "referencia": p.get("referencia", ""),
        "cod_color": p.get("cod_color", ""),
        "tallas": p.get("tallas", ""),
        "n_fotos": len(nombres),
        "repetidas": repetidas,
        "archivos": " ".join(nombres),
        "nota": nota,
        "url_producto": ficha or BUSQUEDA.format(cod=p.get("referencia") or cod),
        "label_producto": "Ver la ficha en harmontblaine.com" if ficha
                          else "Buscar en harmontblaine.com",
        "url_foto1": "",
    }


if __name__ == "__main__":
    main()