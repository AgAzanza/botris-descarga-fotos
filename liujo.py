#!/usr/bin/env python3
"""
LIU JO — baja las fotos de los productos de un excel.


QUE HACE ESTE PROGRAMA, EN CRIOLLO
==================================

Por cada producto del excel arma la direccion de sus fotos y las baja.
No hace falta abrir la ficha ni buscar nada: el nombre de cada imagen es
el codigo del producto mas el numero de foto.

Leer el excel, comparar fotos repetidas y armar la hoja de revision es
igual para todas las marcas y vive en comun.py.


EL CODIGO
---------
Los primeros 16 caracteres son el producto: 11 de referencia y 5 de color.
Lo que sigue es la talla.

    2F3073T0300 61735 TU
    └ referencia ┘└color┘└talla┘

Ver GRAMATICAS_TEXTO en comun.py.


LAS FOTOS
---------
Liu Jo corre sobre Salesforce Commerce Cloud y las fotos estan en su
catalogo con un nombre que se arma solo:

    https://www.liujo.com/on/demandware.static/-/Sites-liujo-master-catalog/
        default/images/E-{codigo de 16}-{numero}-N.jpeg

Verificado a mano:
  - En la web la direccion trae ademas una marca de version (dw2a0d4ee7),
    distinta en cada foto. NO hace falta: sin ella responde igual.
  - LA NUMERACION NO SE PUEDE ADIVINAR. Una camisa iba del 1 al 7; los
    bolsos empiezan en la 7 y saltan (7 y 8, 7 y 12, 7 y 13); un sombrero
    va del 5 al 8. Cada regla que se armo mirando un caso fallo con el
    siguiente. Por eso se prueban TODOS los numeros del 1 al MAX_FOTOS, sin
    cortar en los huecos. Un numero que no existe responde un 404 de unos
    cientos de bytes, asi que probar de mas casi no cuesta, y las pruebas
    van en paralelo.
  - HAY FOTOS DE PRODUCTOS DESPUBLICADOS. Un bolso con la ficha en "not
    found" tenia igual su foto en el catalogo. Por eso no se depende de la
    ficha para saber si hay fotos.
  - LA FICHA SE USA PARA EL ORDEN. Cuando existe, su carrusel dice cual es
    la foto de portada (en el sombrero, la 5; en los bolsos, la 7), que no
    es la del numero mas bajo.
  - Esa direccion entrega la foto ORIGINAL, de alrededor de un mega. Para
    que no pesen tanto se le pide la version redimensionada a la propia
    web (/dw/image/v2/...?sw=2000). Si eso no responde, se baja la
    original y se achica antes de guardarla.

LO QUE HAY QUE SABER
--------------------
Muchos productos del inventario ya no estan publicados: la ficha dice
"sorry, we couldn't find your product". En el caso que se probo, el
producto despublicado tampoco tenia fotos en el catalogo. Asi que es
esperable que una parte grande salga sin fotos. No es un error.


USO
---
    python liujo.py Liujo.xls --limite 30
    python liujo.py Liujo.xls
    python liujo.py Liujo.xls --forzar
"""

import argparse
import re
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pandas as pd
import requests
from requests.adapters import HTTPAdapter

import comun


# ===========================================================================
# LO ESPECIFICO DE LIU JO
# ===========================================================================

MARCA = "LIUJO"
CARPETA = "LiuJo"
SUFIJO = "liujo-ecuador"

CATALOGO = ("on/demandware.static/-/Sites-liujo-master-catalog/default/"
            "images/E-{cod}-{n}-N.jpeg")
ORIGINAL = "https://www.liujo.com/" + CATALOGO
REDIMENSIONADA = "https://www.liujo.com/dw/image/v2/BDNR_PRD/" + CATALOGO + "?sw={ancho}"
FICHA = "https://www.liujo.com/int/{cod}.html"

# Ancho maximo de las fotos guardadas. Las originales de Liu Jo pesan un
# mega cada una; con esto quedan livianas sin perder calidad visible.
MAX_ANCHO = 2000

# Hasta que numero de foto se prueba. El mas alto visto es el 13; se deja
# margen. Si la ficha nombra alguno mas alto, tambien se prueba.
MAX_FOTOS = 20

# Cuantas pruebas a la vez. Las fotos estan en el servidor de archivos
# estaticos de la marca, que aguanta esto sin problema.
HILOS = 8


def url_ficha(cod):
    return FICHA.format(cod=cod)


class Descarga:
    """
    Baja las fotos de un producto.

    Para cada numero prueba primero la version redimensionada, que pesa
    mucho menos. Si no responde pero la original si, pasa a bajar
    originales y achicarlas el resto de la corrida.
    """

    def __init__(self, sesion, ancho=MAX_ANCHO):
        self.sesion = sesion
        self.ancho = ancho
        self.modo = None            # None hasta saberlo, despues "chica" u "original"
        # una conexion por hilo, para que las pruebas en paralelo no esperen
        adaptador = HTTPAdapter(pool_connections=HILOS, pool_maxsize=HILOS)
        sesion.mount("https://", adaptador)

    def foto(self, cod, n):
        if self.modo in (None, "chica"):
            img = comun.bajar(REDIMENSIONADA.format(cod=cod, n=n, ancho=self.ancho),
                              self.sesion)
            if img is not None:
                self.modo = "chica"
                return img
        if self.modo in (None, "original"):
            img = comun.bajar(ORIGINAL.format(cod=cod, n=n), self.sesion)
            if img is not None:
                self.modo = "original"
                if img.width > self.ancho:
                    alto = round(img.height * self.ancho / img.width)
                    img = img.resize((self.ancho, alto))
                return img
        return None

    def orden_de_la_ficha(self, cod):
        """
        El orden del carrusel segun la ficha, o [] si no hay ficha.

        Solo se usa para ordenar: que fotos existen lo dice la prueba de
        todos los numeros, porque hay productos despublicados que igual
        tienen fotos.
        """
        try:
            r = self.sesion.get(FICHA.format(cod=cod), headers=comun.HEADERS,
                                timeout=30, allow_redirects=True)
        except requests.RequestException:
            return []
        if r.status_code != 200:
            return []
        orden = []
        for m in re.finditer(rf"E-{re.escape(cod)}-(\d+)-N", r.text):
            n = int(m.group(1))
            if n not in orden:
                orden.append(n)
        return orden

    def fotos(self, cod):
        """
        Todas las fotos de un producto, en el orden del carrusel.

        Prueba TODOS los numeros del 1 al MAX_FOTOS (y los que nombre la
        ficha), en paralelo y sin cortar en los huecos. Despues ordena:
        primero como las muestra la ficha, y las que la ficha no nombra al
        final, por numero.
        """
        orden = self.orden_de_la_ficha(cod)
        numeros = sorted(set(range(1, MAX_FOTOS + 1)) | set(orden))

        with ThreadPoolExecutor(HILOS) as ex:
            resultados = list(ex.map(lambda n: (n, self.foto(cod, n)), numeros))
        halladas = {n: img for n, img in resultados if img is not None}

        primero = [n for n in orden if n in halladas]
        resto = sorted(n for n in halladas if n not in primero)
        return [(n, halladas[n]) for n in primero + resto]


# ===========================================================================
# PRINCIPAL
# ===========================================================================

def main(argv=None):
    """argv=None usa la linea de comandos; una lista permite llamarlo
    desde codigo (lo hace runner.py cuando corre dentro del .exe)."""
    ap = argparse.ArgumentParser(description="Baja las fotos de Liu Jo.")
    ap.add_argument("excel", type=Path)
    ap.add_argument("--salida", type=Path, default=Path("fotos"))
    ap.add_argument("--limite", type=int, default=None)
    ap.add_argument("--ancho", type=int, default=MAX_ANCHO,
                    help="ancho maximo de las fotos guardadas (por defecto 2000)")
    ap.add_argument("--forzar", action="store_true")
    args = ap.parse_args(argv)

    # referencia, color y talla salen del codigo (ver comun.GRAMATICAS_TEXTO)
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
    print("las direcciones de las fotos se arman desde el codigo\n")

    sesion = requests.Session()
    descarga = Descarga(sesion, args.ancho)
    filas = []

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
            nota = ("El codigo no tiene el formato de Liu Jo (16 caracteres + "
                    "talla). Suele ser una bolsa o un accesorio de empaque.")
        else:
            # Se pasa la POSICION en el carrusel, no el numero de archivo:
            # comun.unicas() ordena por ese valor, y con el numero la portada
            # pasaria a ser la del numero mas bajo en vez de la primera de la
            # ficha (en la camisa de prueba, la 1 en vez de la 4).
            crudas = [(pos, img) for pos, (_, img) in enumerate(descarga.fotos(cod))]
            if not crudas:
                nota = ("Liu Jo no tiene fotos de este producto en su catalogo. "
                        "Lo mas probable es que ya no este publicado.")

        imagenes = comun.unicas(crudas)
        nombres = comun.guardar(imagenes, carpeta, cod, SUFIJO)
        repetidas = len(crudas) - len(imagenes)

        detalle = f"{len(nombres)} fotos" if nombres else (nota or "sin fotos")
        if repetidas:
            detalle += f"  ({repetidas} repetidas descartadas)"
        print(f"  {i:>4}/{len(prod)} {'[ok]' if nombres else '[!!]'} {cod:<18} {detalle[:70]}")

        filas.append(_fila(p, nombres, repetidas, nota))

    rep = pd.DataFrame(filas)
    destino, total, sin = comun.hoja_revision(rep, carpeta, CARPETA)
    xls, n_filas, n_prod = comun.excel_encontrados(args.excel, rep, carpeta,
                                                   marca=MARCA, gram=gram)

    print(f"\n{'-' * 64}")
    print(f"  productos con fotos  : {(rep.n_fotos > 0).sum()} de {len(rep)}")
    print(f"  fotos guardadas      : {total}")
    print(f"  repetidas descartadas: {rep.repetidas.sum()}")
    if descarga.modo:
        print(f"  tamano pedido        : "
              + ("redimensionadas por la web" if descarga.modo == "chica"
                 else "originales, achicadas al guardar"))
    if xls:
        print(f"  excel de encontrados : {xls.name}  "
              f"({n_prod} productos, {n_filas} filas)")
    print(f"\n  Abre:  {destino.resolve()}")


def _fila(p, nombres, repetidas, nota=""):
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
        "url_producto": url_ficha(cod),
        "label_producto": "Ver la ficha en liujo.com",
        "url_foto1": ORIGINAL.format(cod=cod, n=1) if nombres else "",
    }


if __name__ == "__main__":
    main()