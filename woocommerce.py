#!/usr/bin/env python3
"""
woocommerce.py — arma el archivo de carga de WooCommerce.

Toma el excel de una marca y las fotos que ya se bajaron, y produce el
xlsx listo para subir.


COMO QUEDA EL ARCHIVO
=====================
Una fila PADRE por referencia, con todas las tallas y colores juntos,
seguida de una fila por cada variacion:

  ID         SKU                  Type      Parent       Talla              Color
  300036789  D650ZB00022          variable               35|36|36.5|37|...  Azul|Crema
  300036790  D650ZB00022C400535   variable  D650ZB00022  35                 Azul
  300036791  D650ZB00022C400536   variable  D650ZB00022  36                 Azul

El SKU de cada variacion es el CODIGO original tal cual viene del excel,
con la talla y con sus puntos si los tiene (Lanidor: "405608.188.L"). El
del padre es la referencia.

Las listas van separadas con "|". Los productos de una sola fila salen
como "simple", sin padre.


LAS IMAGENES
============
Aca esta la diferencia mas grande con el plugin de PHP que hacia esto
antes. Ese generaba las URLs por convencion y asumia SIEMPRE 3 fotos por
producto:

    {base}/{codigo}-{sufijo}.webp
    {base}/{codigo}-1-{sufijo}.webp
    {base}/{codigo}-2-{sufijo}.webp

Si el producto tenia 2 fotos, WooCommerce quedaba con una imagen rota. Si
tenia 8, se perdian 5.

Nosotros ya bajamos las fotos, asi que sabemos cuantas hay y como se
llaman: las URLs salen de los archivos que existen en disco. Ni una de
mas ni una de menos.


LOS IDS
=======
WooCommerce necesita un ID unico por fila y no se pueden repetir JAMAS
entre corridas. Se llevan en ids.json, que se puede poner en una carpeta
compartida para que dos computadoras usen el mismo contador.


USO
---
    python woocommerce.py entradas/Geox.xls --marca GEOX
    python woocommerce.py entradas/Geox.xls --marca GEOX --fotos fotos
    python woocommerce.py entradas/Geox.xls --marca GEOX --sin-ia
"""

import argparse
import json
import re
from datetime import datetime
from pathlib import Path

import pandas as pd

import comun

try:
    import claude
    HAY_CLAUDE = True
except ImportError:
    HAY_CLAUDE = False


# ===========================================================================
# CONFIGURACION
# ===========================================================================

COLUMNAS = [
    "ID", "Title", "content", "SKU", "Price", "Product Type",
    "Parent Product ID", "Fit Del Producto", "Color del producto",
    "Talla Del Producto", "Material Del Producto", "Marcas", "Image URL",
    "Categoria",
]

SEP = "|"
BASE_URL = "https://www.botris.com.ec/e-comm/Boss/2025-07"
# Desde donde arranca el contador la PRIMERA vez, si no existe ids.json.
#
# Sale de los archivos que ya se subieron: Geox uso 300036789..300037032 y
# Lanidor siguio con 300037033..300037319. El ultimo usado es 300037319,
# asi que el primero libre es el siguiente.
#
# Una vez que ids.json existe, este numero no se mira mas. Para moverlo a
# mano: --desde-id
ID_INICIAL = 300037320

# Nombre visible de la marca y carpeta donde estan sus fotos.
MARCAS = {
    "GEOX":   {"nombre": "Geox",          "carpeta": "Geox",        "sufijo": "geox-ecuador"},
    "LANI":   {"nombre": "Lanidor",       "carpeta": "Lanidor",     "sufijo": "lanidor-ecuador"},
    "FUSTER": {"nombre": "Fuster",        "carpeta": "Fuster",      "sufijo": "fuster-ecuador"},
    "HUGO":   {"nombre": "Hugo",          "carpeta": "HUGO",        "sufijo": "hugo-ecuador"},
    "BOSS":   {"nombre": "Boss",          "carpeta": "BOSS",        "sufijo": "boss-ecuador"},
    "MK":     {"nombre": "Michael Kors",  "carpeta": "MichaelKors", "sufijo": "michaelkors-ecuador"},
}

# Categoria por la primera letra del codigo. Es la nomenclatura de Geox;
# las demas marcas no la usan y caen en el respaldo por tipo de prenda.
CAT_LETRA = {"D": "Mujer", "U": "Hombre", "B": "Ninos", "J": "Ninos", "K": "Ninos"}
TIPOS_MUJER = {
    "SKIRT", "DRESS", "BLOUSE", "SANDALS", "BALLERINA", "HANDBAG", "BAG",
    "PONCHO", "KIMONO", "PAREO", "SCARF", "SQUARE", "NECKLACE", "EARRINGS",
    "BRACELET", "RING", "BROOCH", "ELASTIC", "TUNIC",
}


# ===========================================================================
# CONTADOR DE IDS
# ===========================================================================

class Contador:
    """
    Reparte IDs sin repetir nunca.

    Se guarda en disco ANTES de usar el bloque, no despues: si el programa
    se cae a la mitad, esos IDs quedan quemados y no se reparten otra vez.
    Perder numeros no cuesta nada; repetirlos rompe la carga.
    """

    def __init__(self, archivo=None):
        self.archivo = Path(archivo) if archivo else Path(__file__).parent / "ids.json"
        self.proximo = ID_INICIAL
        if self.archivo.exists():
            try:
                self.proximo = int(json.loads(
                    self.archivo.read_text(encoding="utf-8"))["proximo"])
            except Exception:
                pass

    def fijar(self, numero):
        """Mueve el contador a mano, sin generar nada."""
        self.proximo = int(numero)
        self.archivo.write_text(json.dumps({
            "proximo": self.proximo,
            "ultimo_uso": datetime.now().isoformat(timespec="seconds"),
        }, indent=1), encoding="utf-8")
        return self.proximo

    def reservar(self, cuantos):
        """Aparta un bloque y lo deja anotado. Devuelve el primer ID."""
        desde = self.proximo
        self.proximo = desde + cuantos
        self.archivo.write_text(json.dumps({
            "proximo": self.proximo,
            "ultimo_uso": datetime.now().isoformat(timespec="seconds"),
        }, indent=1), encoding="utf-8")
        return desde


# ===========================================================================
# HELPERS DE FORMATO
# ===========================================================================

def titulo(texto):
    """'D XAND 3' -> 'D Xand 3'   |   'HIGH-LOW' -> 'High-Low'"""
    t = re.sub(r"\s+", " ", str(texto)).strip().lower()
    return " ".join("-".join(p[:1].upper() + p[1:] for p in palabra.split("-"))
                    for palabra in t.split(" "))


def precio(valor):
    """'169.00' -> '169'   |   '149.90' -> '149.9'   |   vacio -> ''"""
    try:
        n = float(str(valor).replace(",", ".").strip())
    except (ValueError, AttributeError):
        return ""
    if n <= 0:
        return ""
    return str(int(n)) if n == int(n) else f"{n:g}"


def talla(valor):
    """'36.0' -> '36'   |   '36.5' -> '36.5'   |   'O/S' -> 'O/S'"""
    t = str(valor).replace(",", ".").strip()
    if re.fullmatch(r"\d+\.0+", t):
        return str(int(float(t)))
    return t


def categoria(cod, tipo=""):
    letra = str(cod)[:1].upper()
    if letra in CAT_LETRA:
        return CAT_LETRA[letra]
    if str(tipo).strip().upper() in TIPOS_MUJER:
        return "Mujer"
    return ""


def sin_repetir(valores):
    return list(dict.fromkeys(v for v in valores if str(v).strip()))


# ===========================================================================
# LECTURA
# ===========================================================================

def _columnas(df, usar_ia, clave=None):
    """
    De que columna sale cada dato.

    Primero los alias fijos de comun.py, que cubren los archivos conocidos y
    no cuestan nada. Si falta alguno importante, se le pregunta a Claude:
    las marcas cambian los encabezados entre envios y ninguna lista fija
    aguanta eso para siempre.
    """
    mapa = {}
    for canonico in ("codigo", "referencia", "nombre", "talla", "color", "marca"):
        real = comun._columna_real(df, canonico)
        if real:
            mapa[canonico] = real

    # Los alias de comun.py no cubren precio, material, fit ni tipo: se
    # buscan por patron en el nombre de la columna.
    #
    # OJO CON EL ORDEN Y CON LAS COLUMNAS YA USADAS: en el excel de Hugo la
    # columna que se llama MATERIAL no es el material, es la REFERENCIA del
    # producto, y el material real esta en ART_COMPO_VESTI. Por eso se
    # prueban primero los patrones inequivocos (COMPO, VESTI) y se saltean
    # las columnas que ya tienen otro rol asignado. Sin eso, la composicion
    # de la prenda se llenaria con codigos de producto.
    patrones = (
        ("precio",   r"PREC|PRICE|PVP"),
        ("material", r"COMPO|VESTI|TEJID|COMPOSIC|MATER"),
        ("fit",      r"\bFIT\b|AJUST|CORTE"),
        ("tipo",     r"GRUPO|TIPO|FAMIL"),
    )
    for canonico, patron in patrones:
        if canonico in mapa:
            continue
        for c in df.columns:
            if c in mapa.values():
                continue
            if re.search(patron, str(c).upper()):
                mapa[canonico] = c
                break

    faltan = [c for c in ("codigo", "referencia", "nombre", "talla") if c not in mapa]
    if faltan and usar_ia and HAY_CLAUDE:
        print(f"  faltan columnas {faltan}: preguntando a Claude...")
        filas = df.head(claude.FILAS_MUESTRA).to_dict("records")
        det = claude.detectar_columnas(filas, clave)
        for canonico in ("sku", "referencia", "nombre", "precio", "talla",
                         "color", "material", "fit", "tipo", "marca"):
            valor = det.get(f"col_{canonico}")
            destino = "codigo" if canonico == "sku" else canonico
            if valor and valor in df.columns and destino not in mapa:
                mapa[destino] = valor
    return mapa


# ===========================================================================
# ARMADO
# ===========================================================================

def _fila(id_, sku, nombre, precio_, tipo, padre, color, talla_,
          material, fit, marca, urls, cat):
    return {
        "ID": str(id_), "Title": nombre, "content": "", "SKU": sku,
        "Price": precio_, "Product Type": tipo, "Parent Product ID": padre,
        "Fit Del Producto": fit, "Color del producto": color,
        "Talla Del Producto": talla_, "Material Del Producto": material,
        "Marcas": marca, "Image URL": urls, "Categoria": cat,
    }


def construir(excel, marca, carpeta_fotos, base_url=BASE_URL,
              usar_ia=True, clave=None, contador=None):
    """
    Devuelve (filas, avisos, resumen).

    Las filas del excel se agrupan por REFERENCIA. Dentro de cada grupo,
    cada fila del excel es una variacion (una talla de un color).
    """
    cfg = MARCAS.get(marca.upper(), {})
    sufijo = cfg.get("sufijo", f"{marca.lower()}-ecuador")
    nombre_marca = cfg.get("nombre", marca.title())

    crudo = comun.leer(excel)
    cols = _columnas(crudo, usar_ia, clave)

    col_marca = cols.get("marca")
    if col_marca:
        crudo = crudo[crudo[col_marca].astype(str).str.strip().str.upper()
                      == marca.upper()]
    if len(crudo) == 0:
        return [], [f"El excel no tiene filas de {marca}."], {}

    c_cod, c_ref = cols.get("codigo"), cols.get("referencia")
    c_nom, c_tal = cols.get("nombre"), cols.get("talla")
    c_pre, c_col = cols.get("precio"), cols.get("color")
    c_mat, c_fit, c_tip = cols.get("material"), cols.get("fit"), cols.get("tipo")

    # --- traducciones, una sola vez para todo el archivo ---
    trad_color, trad_mat = {}, {}
    if usar_ia and HAY_CLAUDE:
        claude.reiniciar_uso()
        try:
            if c_col:
                trad_color = claude.traducir_colores(crudo[c_col].tolist(), clave)
            if c_mat:
                trad_mat = claude.traducir_materiales(crudo[c_mat].tolist(), clave)
        except claude.SinClave as e:
            print(f"  {e}\n  Sigo sin traducir.")
        except Exception as e:
            print(f"  Claude fallo ({str(e)[:70]}). Sigo sin traducir.")

    def color_de(v):
        v = str(v).strip()
        return trad_color.get(v, v.capitalize()) if v else ""

    def material_de(v):
        v = str(v).strip()
        if not v:
            return ""
        if v in trad_mat:
            return trad_mat[v]
        return claude.material_dominante(v) if HAY_CLAUDE else v.capitalize()

    # --- agrupar por referencia, conservando el orden del excel ---
    grupos, orden = {}, []
    for _, r in crudo.iterrows():
        ref = str(r[c_ref]).strip() if c_ref else ""
        if not ref:
            continue
        if ref not in grupos:
            grupos[ref] = []
            orden.append(ref)
        grupos[ref].append(r)

    # cuantas filas vamos a generar, para reservar los IDs de una vez
    total = sum(len(g) + (1 if len(g) > 1 else 0) for g in grupos.values())
    contador = contador or Contador()
    id_ = contador.reservar(total)
    desde = id_

    filas, avisos = [], []
    n_padres = n_var = n_simple = 0
    sin_foto = []

    for ref in orden:
        grupo = grupos[ref]

        variaciones = []
        for r in grupo:
            cod_prod = comun.codigo_producto(r[c_cod], r[c_tal] if c_tal else "")
            archivos = comun.archivos_de(carpeta_fotos, cod_prod, sufijo)
            if not archivos:
                sin_foto.append(cod_prod)
            variaciones.append({
                "sku": str(r[c_cod]).strip(),
                "cod": cod_prod,
                "talla": talla(r[c_tal]) if c_tal else "",
                "color": color_de(r[c_col]) if c_col else "",
                "material": material_de(r[c_mat]) if c_mat else "",
                "fit": str(r[c_fit]).strip().capitalize() if c_fit else "",
                "nombre": titulo(r[c_nom]) if c_nom else "",
                "precio": precio(r[c_pre]) if c_pre else "",
                "urls": SEP.join(f"{base_url.rstrip('/')}/{a}" for a in archivos),
                "cat": categoria(cod_prod, r[c_tip] if c_tip else ""),
            })

        # los que no tienen ninguna foto no van al archivo de carga
        variaciones = [v for v in variaciones if v["urls"]]
        if not variaciones:
            continue

        if len(variaciones) == 1:
            v = variaciones[0]
            filas.append(_fila(id_, v["sku"], v["nombre"], v["precio"], "simple",
                               "", v["color"], v["talla"], v["material"],
                               v["fit"], nombre_marca, v["urls"], v["cat"]))
            id_ += 1
            n_simple += 1
            continue

        p = variaciones[0]
        colores = sin_repetir(v["color"] for v in variaciones)
        tallas = sin_repetir(v["talla"] for v in variaciones)

        # EL COLOR, EN EL PADRE Y EN LOS HIJOS, SIGUEN REGLAS DISTINTAS.
        #
        # PADRE: la lista de colores solo si hay DOS O MAS. Con uno solo va
        #   en blanco, porque el padre agrupa por referencia y ese dato no
        #   distingue nada. Es la regla que ya usaba el sistema anterior.
        #
        # HIJOS: su color SIEMPRE, aunque todas las variaciones del grupo
        #   sean del mismo. Aca el sistema anterior se equivocaba: aplicaba
        #   la condicion del padre tambien a los hijos, y como la mayoria de
        #   las referencias tienen un unico color (23 de 25 en Hugo, 27 de
        #   33 en Geox), las variaciones quedaban sin color casi siempre. Su
        #   archivo de Geox salio con CERO colores en 244 filas.
        color_padre = SEP.join(colores) if len(colores) > 1 else ""

        filas.append(_fila(id_, ref, p["nombre"], p["precio"], "variable", "",
                           color_padre, SEP.join(tallas), p["material"],
                           p["fit"], nombre_marca, p["urls"], p["cat"]))
        id_ += 1
        n_padres += 1

        for v in variaciones:
            filas.append(_fila(id_, v["sku"], v["nombre"], v["precio"],
                               "variable", ref, v["color"],
                               v["talla"], v["material"], v["fit"],
                               nombre_marca, v["urls"], v["cat"]))
            id_ += 1
            n_var += 1

    if sin_foto:
        avisos.append(f"{len(sin_foto)} productos sin fotos quedaron fuera "
                      f"(ej: {', '.join(sin_foto[:3])}).")

    resumen = {"total": len(filas), "padres": n_padres, "variaciones": n_var,
               "simples": n_simple, "id_desde": desde, "id_hasta": id_ - 1,
               "sin_foto": len(sin_foto),
               "con_color": sum(1 for f in filas if f["Color del producto"]),
               "hijos": sum(1 for f in filas if f["Parent Product ID"])}
    if usar_ia and HAY_CLAUDE:
        resumen["claude"] = dict(claude.ULTIMO)
    return filas, avisos, resumen


# ===========================================================================
# VALIDACION
# ===========================================================================

def validar(filas):
    """
    Revisa el archivo ANTES de escribirlo.

    Existe por un error real encontrado en la salida del sistema anterior:
    cinco productos de Geox tenian quince variaciones con las tallas
    repetidas tres veces (tres colores del mismo modelo) y el atributo de
    color vacio. WooCommerce no puede distinguirlas: son quince variaciones
    con la misma combinacion de atributos. El archivo se genero igual, con
    estadisticas que decian que estaba todo bien.

    Devuelve una lista de problemas, del mas grave al menos.
    """
    problemas = []
    porpadre = {}
    for f in filas:
        if f["Parent Product ID"]:
            porpadre.setdefault(f["Parent Product ID"], []).append(f)

    for padre, hijos in porpadre.items():
        combos = [(h["Talla Del Producto"], h["Color del producto"]) for h in hijos]
        repes = len(combos) - len(set(combos))
        if repes:
            problemas.append(
                f"GRAVE  {padre}: {len(hijos)} variaciones con {repes} "
                f"combinaciones talla+color repetidas. WooCommerce no las "
                f"puede distinguir. Suele ser que falto detectar el color.")

    skus = [f["SKU"] for f in filas]
    for s in {x for x in skus if skus.count(x) > 1}:
        problemas.append(f"GRAVE  SKU duplicado: {s}")

    ids = [f["ID"] for f in filas]
    if len(ids) != len(set(ids)):
        problemas.append("GRAVE  hay IDs repetidos")

    sin_precio = [f["SKU"] for f in filas if not f["Price"]]
    if sin_precio:
        problemas.append(f"aviso  {len(sin_precio)} filas sin precio "
                         f"(ej: {', '.join(sin_precio[:3])})")

    sin_cat = [f["SKU"] for f in filas if not f["Categoria"]]
    if sin_cat:
        problemas.append(f"aviso  {len(sin_cat)} filas sin categoria")

    return problemas


# ===========================================================================
# ESCRITURA
# ===========================================================================

def escribir(filas, destino):
    """
    Escribe el CSV de carga.

    En CSV y no en xlsx porque es lo que espera el importador, y con el
    mismo formato que los archivos que hoy se suben: sin BOM, separador
    coma, fin de linea CRLF y UTF-8.
    """
    df = pd.DataFrame(filas, columns=COLUMNAS)
    destino = Path(destino)
    destino.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(destino, index=False, encoding="utf-8", lineterminator="\r\n")
    return destino


# ===========================================================================
# PRINCIPAL
# ===========================================================================

def carpeta_de_fotos(excel, indicada, carpeta_marca):
    """
    Decide donde estan las fotos, probando lo mas probable primero.

    1. AL LADO DEL EXCEL. Es el caso normal: la descarga deja el
       encontrados.xlsx junto a la carpeta fotos/, asi que si el excel que
       nos pasan tiene una carpeta fotos/ al lado, esa es la carpeta de la
       marca. Se prueba primero porque es lo que hace la gente: pasarle el
       encontrados.xlsx y esperar que se entienda solo.
    2. Lo que diga --fotos, mas la carpeta de la marca: fotos/HUGO
    3. Lo que diga --fotos, si ya apunta a la carpeta de una marca.

    Devuelve (carpeta, donde_se_busco) para poder explicar si no aparece.
    """
    intentos = []

    padre = Path(excel).resolve().parent
    intentos.append(padre)
    if (padre / comun.CARPETA_FOTOS).is_dir():
        return padre, intentos

    if indicada:
        con_marca = Path(indicada) / carpeta_marca
        intentos.append(con_marca)
        if (con_marca / comun.CARPETA_FOTOS).is_dir():
            return con_marca, intentos

        intentos.append(Path(indicada))
        if (Path(indicada) / comun.CARPETA_FOTOS).is_dir():
            return Path(indicada), intentos

    return None, intentos


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Arma el archivo de carga de WooCommerce.")
    ap.add_argument("excel", type=Path,
                    help="el excel original de la marca, o el encontrados.xlsx "
                         "que dejo la descarga (tiene las mismas columnas)")
    ap.add_argument("--marca", default=None, help="GEOX, LANI, HUGO, BOSS, MK...")
    ap.add_argument("--fotos", type=Path, default=Path("fotos"),
                    help="carpeta de fotos. Por defecto se usa la que esta al "
                         "lado del excel, que es donde la deja la descarga")
    ap.add_argument("--salida", type=Path, default=None)
    ap.add_argument("--base-url", default=BASE_URL)
    ap.add_argument("--ids", default=None, help="archivo del contador de IDs")
    ap.add_argument("--desde-id", type=int, default=None,
                    help="fija desde que numero siguen los IDs y termina")
    ap.add_argument("--sin-ia", action="store_true",
                    help="no consultar a Claude (no traduce colores)")
    ap.add_argument("--renombrar-fotos", action="store_true",
                    help="le agrega el sufijo de marca a las fotos que no lo "
                         "tengan (las bajadas antes de que el sufijo existiera)")
    args = ap.parse_args(argv)
    if not args.marca and not args.desde_id:
        ap.error("falta --marca")

    if args.desde_id:
        c = Contador(args.ids)
        print(f"\n  contador movido a {c.fijar(args.desde_id)}  ({c.archivo})")
        return

    marca = args.marca.upper()
    cfg = MARCAS.get(marca, {})

    print(f"\n{marca}: {args.excel.name}")
    carpeta, buscado = carpeta_de_fotos(args.excel, args.fotos,
                                        cfg.get("carpeta", marca))
    if carpeta is None:
        print(f"\n  No encontre la carpeta con las fotos. Busque en:")
        for d in buscado:
            print(f"    {d}/{comun.CARPETA_FOTOS}")
        print(f"\n  Indicala con --fotos, apuntando a la carpeta de la marca")
        print(f"  (la que contiene '{comun.CARPETA_FOTOS}' adentro).")
        return
    print(f"  fotos en: {carpeta / comun.CARPETA_FOTOS}")

    if args.renombrar_fotos:
        hechas, ya = comun.renombrar_con_sufijo(carpeta, cfg.get("sufijo", ""))
        print(f"  renombradas: {hechas}  (ya tenian el sufijo: {ya})")

    filas, avisos, resumen = construir(
        args.excel, marca, carpeta, args.base_url,
        usar_ia=not args.sin_ia, contador=Contador(args.ids))

    if not filas:
        print("\n  No se genero ninguna fila.")
        for a in avisos:
            print(f"  {a}")
        return

    problemas = validar(filas)
    graves = [p for p in problemas if p.startswith("GRAVE")]

    # --- que hizo Claude, para poder comprobar que funciona ---
    if args.sin_ia:
        print("  Claude         : desactivado con --sin-ia")
    elif not HAY_CLAUDE:
        print("  Claude         : falta el archivo claude.py")
    else:
        u = resumen.get("claude", {})
        if not claude.leer_clave():
            print("  Claude         : SIN CLAVE. Los colores salen sin traducir.")
            print("                   Copiá .env.ejemplo como .env y poné tu clave.")
        else:
            print(f"  Claude         : {u.get('llamadas', 0)} llamadas · "
                  f"{u.get('nuevos', 0)} valores nuevos · "
                  f"{u.get('del_cache', 0)} del cache")
            print(f"                   columnas: {u.get('deteccion') or 'no hizo falta'}"
                  f"  ({claude.MODELO_DETECCION} / {claude.MODELO_TRADUCCION})")

    print(f"\n  filas          : {resumen['total']}")
    print(f"  padres         : {resumen['padres']}")
    print(f"  variaciones    : {resumen['variaciones']}")
    print(f"  simples        : {resumen['simples']}")
    print(f"  con color      : {resumen['con_color']}  "
          f"(variaciones: {resumen['hijos']}, simples: {resumen['simples']}; "
          f"el padre solo si tiene 2 colores o mas)")
    print(f"  IDs            : {resumen['id_desde']} a {resumen['id_hasta']}")
    for a in avisos:
        print(f"  {a}")

    if problemas:
        print(f"\n  revision:")
        for p in problemas:
            print(f"    {p}")

    if graves:
        print(f"\n  {len(graves)} problemas GRAVES. No escribo el archivo: "
              f"subirlo asi rompe productos en la tienda.")
        print("  Corregilos y volve a correr.")
        return

    salida = args.salida or (carpeta / f"woocommerce_{cfg.get('sufijo', marca.lower())}"
                             f"_{datetime.now():%Y%m%d_%H%M%S}.csv")
    escribir(filas, salida)
    print(f"\n  Listo: {Path(salida).resolve()}")


if __name__ == "__main__":
    main()