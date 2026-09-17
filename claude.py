#!/usr/bin/env python3
"""
claude.py — le pregunta a Claude lo que no se puede programar con reglas.

POR QUE HACE FALTA LA IA ACA
============================
Las marcas no mandan siempre el mismo excel. Cambian los titulos de las
columnas entre envios, agregan columnas, traducen encabezados. Una lista
fija de alias funciona hasta que llega el archivo que no encaja, y ahi hay
que parchear el codigo otra vez.

Ademas los colores y materiales vienen en el idioma de cada marca
(portugues en Lanidor, ingles en MK, italiano en algunos) y con formatos
como "003% Spandex; 097% Cotone;". Traducir eso con un diccionario es una
pelea sin fin.

QUE SE HACE PARA QUE SALGA BARATO
=================================
1. UN MODELO POR TAREA: Sonnet para decidir que significa cada columna
   (se hace UNA vez por formato y queda cacheado; equivocarse ahi arruina
   el archivo entero) y Haiku, mucho mas barato, para las traducciones
   (mucho volumen, riesgo nulo, error visible a simple vista).

2. CACHE EN DISCO: la respuesta se guarda contra una firma del pedido. El
   archivo de Geox tiene siempre las mismas columnas, asi que se consulta
   UNA vez y las siguientes veinte corridas no gastan nada. Los colores se
   cachean de a uno, asi que un archivo nuevo de la misma marca solo paga
   por los colores que no vio antes.

3. SOLO LO UNICO: no se mandan filas enteras sino 3 valores de ejemplo por
   columna, recortados. Con 58 columnas eso baja el pedido a menos de un
   tercio. De colores y materiales, solo los valores distintos: en los
   cinco archivos de prueba son 138 colores y 234 materiales en total.

4. POR LOTES: las traducciones van de a 120 por llamada. El plugin de PHP
   pedia todo junto con un limite de 2.000 tokens y con archivos grandes
   la respuesta se cortaba a la mitad, dejando un JSON roto.

LA CLAVE
========
Se lee de la variable de entorno ANTHROPIC_API_KEY o de un archivo .env
al lado del programa:

    ANTHROPIC_API_KEY=sk-ant-...

Ese archivo NO va al repositorio y NO puede quedar adentro del .exe: se
deja al lado del ejecutable y cada maquina pone la suya.
"""

import hashlib
import json
import os
import re
from pathlib import Path

import requests


# DOS MODELOS, SEGUN LA TAREA
#
# Detectar que significa cada columna es de ALTO RIESGO y BAJISIMO VOLUMEN:
# se hace una vez por formato de archivo y queda cacheado para siempre. Y
# equivocarse sale caro. Los excels traen columnas como CODIGO, CODIGO
# PADRE, REFERENCIA, MATERIAL, CTA_CODIGO_INV, CTA_CODIGO_GST... y en el
# archivo de Hugo la columna MATERIAL no es el material, es la referencia.
# Elegir mal ahi genera SKUs incorrectos en todo el archivo, en silencio.
# Por eso va Sonnet.
#
# Traducir colores y materiales es lo contrario: mucho volumen, riesgo nulo
# y un error se ve a simple vista en la hoja de revision. Por eso va Haiku,
# que es varias veces mas barato.
MODELO_DETECCION = "claude-sonnet-4-6"
MODELO_TRADUCCION = "claude-haiku-4-5-20251001"

API = "https://api.anthropic.com/v1/messages"
VERSION_API = "2023-06-01"

VALORES_MUESTRA = 3      # cuantos valores de ejemplo por columna
LARGO_MUESTRA = 28       # a cuantos caracteres se recorta cada valor
FILAS_MUESTRA = 40       # de cuantas filas se sacan esos valores
LOTE = 120               # cuantos colores o materiales por llamada

AQUI = Path(__file__).parent
CACHE = AQUI / ".cache-claude"

# Que hizo Claude en esta corrida. Lo lee woocommerce.py para poder
# informarlo: sin esto no hay forma de saber si tradujo de verdad, si uso el
# cache o si directamente no se llamo.
ULTIMO = {"llamadas": 0, "nuevos": 0, "del_cache": 0, "deteccion": None}


def reiniciar_uso():
    ULTIMO.update({"llamadas": 0, "nuevos": 0, "del_cache": 0, "deteccion": None})


# ===========================================================================
# CLAVE
# ===========================================================================

def leer_clave():
    """De la variable de entorno, o del .env que este al lado."""
    clave = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if clave:
        return clave
    for carpeta in (AQUI, Path.cwd()):
        env = carpeta / ".env"
        if env.exists():
            for linea in env.read_text(encoding="utf-8").splitlines():
                if linea.strip().startswith("ANTHROPIC_API_KEY"):
                    return linea.split("=", 1)[-1].strip().strip('"').strip("'")
    return ""


class SinClave(RuntimeError):
    pass


# ===========================================================================
# CACHE
# ===========================================================================

def _firma(*partes):
    crudo = json.dumps(partes, sort_keys=True, ensure_ascii=False)
    return hashlib.sha1(crudo.encode("utf-8")).hexdigest()[:16]


def _cache_leer(nombre):
    f = CACHE / f"{nombre}.json"
    if f.exists():
        try:
            return json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            pass
    return None


def _cache_escribir(nombre, dato):
    CACHE.mkdir(exist_ok=True)
    (CACHE / f"{nombre}.json").write_text(
        json.dumps(dato, ensure_ascii=False, indent=1), encoding="utf-8")


# ===========================================================================
# LLAMADA
# ===========================================================================

def _pedir(prompt, max_tokens=2000, clave=None, sistema=None,
           modelo=MODELO_TRADUCCION):
    clave = clave or leer_clave()
    if not clave:
        raise SinClave(
            "Falta la clave de Claude. Poné un archivo .env al lado del "
            "programa con:\n    ANTHROPIC_API_KEY=sk-ant-...")

    cuerpo = {
        "model": modelo,
        "max_tokens": max_tokens,
        "messages": [{"role": "user", "content": prompt}],
    }
    if sistema:
        cuerpo["system"] = sistema

    r = requests.post(API, timeout=90, headers={
        "content-type": "application/json",
        "x-api-key": clave,
        "anthropic-version": VERSION_API,
    }, json=cuerpo)

    if r.status_code != 200:
        detalle = ""
        try:
            detalle = r.json().get("error", {}).get("message", "")
        except Exception:
            detalle = r.text[:200]
        raise RuntimeError(f"Claude respondio {r.status_code}: {detalle}")

    datos = r.json()
    texto = "".join(b.get("text", "") for b in datos.get("content", []))
    uso = datos.get("usage", {})
    return texto, uso


def _json_de(texto):
    """Saca el JSON de la respuesta, tolerando ```json y comentarios //."""
    t = texto.strip()
    t = re.sub(r"^```(?:json)?\s*", "", t)
    t = re.sub(r"\s*```$", "", t)
    t = re.sub(r"//[^\n]*", "", t)          # el modelo a veces comenta
    ini, fin = t.find("{"), t.rfind("}")
    if ini >= 0 and fin > ini:
        t = t[ini:fin + 1]
    return json.loads(t)


# ===========================================================================
# 1. QUE SIGNIFICA CADA COLUMNA
# ===========================================================================

SISTEMA = ("Sos un asistente que clasifica datos de archivos de productos de "
           "moda. Respondes SOLO con JSON valido, sin markdown ni comentarios.")


def _muestra_por_columna(filas):
    """
    {columna: [3 valores distintos]} en vez de filas enteras.

    Mandar 8 filas de 58 columnas son ~2.800 tokens y la mitad son celdas
    vacias o repetidas. Asi baja a menos de un tercio y ademas le sirve mas
    al modelo: lo que necesita para decidir que es cada columna son ejemplos
    de sus VALORES, no filas completas.
    """
    salida = {}
    for col in filas[0].keys():
        vistos = []
        for fila in filas[:FILAS_MUESTRA]:
            v = str(fila.get(col, "")).strip()[:LARGO_MUESTRA]
            if v and v not in vistos:
                vistos.append(v)
            if len(vistos) >= VALORES_MUESTRA:
                break
        salida[col] = vistos
    return salida


def detectar_columnas(filas, clave=None, forzar=False):
    """
    Dice que columna es el SKU, la referencia, el precio, la talla, etc.

    Se cachea contra la lista de encabezados: el mismo formato de archivo no
    se vuelve a consultar nunca.
    """
    if not filas:
        return {}
    encabezados = list(filas[0].keys())

    nombre = "cols-" + _firma(encabezados)
    if not forzar:
        guardado = _cache_leer(nombre)
        if guardado:
            ULTIMO["deteccion"] = "del cache"
            return guardado

    muestra = _muestra_por_columna(filas)
    prompt = (
        "Analiza las columnas de un excel de productos de moda y decime que "
        "significa cada una. Para cada columna te doy algunos valores de "
        "ejemplo.\n\n"
        f"COLUMNAS Y EJEMPLOS:\n{json.dumps(muestra, ensure_ascii=False, indent=0)}\n\n"
        "CUIDADO: hay columnas con nombres parecidos que son cosas distintas "
        "(CODIGO vs CODIGO PADRE vs REFERENCIA vs CTA_CODIGO_*), y nombres "
        "que enganan: en algunos archivos la columna llamada MATERIAL no es "
        "el material sino la referencia del producto. Guiate por los VALORES, "
        "no por el nombre.\n\n"
        "Devolve exactamente este JSON, con el NOMBRE EXACTO de la columna "
        "o null si no existe:\n"
        "{\n"
        '  "col_sku": "codigo unico de cada variacion (incluye la talla)",\n'
        '  "col_referencia": "codigo padre que agrupa las variaciones",\n'
        '  "col_nombre": "nombre del producto",\n'
        '  "col_precio": "precio de venta, con valores mayores a cero",\n'
        '  "col_talla": "talla",\n'
        '  "col_color": "color",\n'
        '  "col_material": "material o composicion, o null",\n'
        '  "col_fit": "fit, corte o ajuste, o null",\n'
        '  "col_tipo": "tipo de prenda (SKIRT, DRESS, SHOES...), o null",\n'
        '  "col_marca": "marca, o null"\n'
        "}"
    )
    texto, uso = _pedir(prompt, 700, clave, SISTEMA, MODELO_DETECCION)
    ULTIMO["llamadas"] += 1
    ULTIMO["deteccion"] = "consultada"
    datos = _json_de(texto)
    datos["_tokens"] = [uso.get("input_tokens"), uso.get("output_tokens")]
    _cache_escribir(nombre, datos)
    return datos


# ===========================================================================
# 2. TRADUCIR COLORES Y MATERIALES
# ===========================================================================

REGLAS_COLOR = (
    "- Traducir al espanol venga del idioma que venga\n"
    "- Nombres SIMPLES: 'Azul marino' -> 'Azul', 'Verde oscuro' -> 'Verde'\n"
    "- Combinaciones con /: 'NAVY/RED' -> 'Azul/Rojo'\n"
    "- Patrones: PRINTED->Estampado, STRIPES->Rayas, FLORAL->Floral, "
    "LEOPARD->Leopardo\n"
    "- Si ya esta en espanol, dejarlo con mayuscula inicial"
)

REGLAS_MATERIAL = (
    "- Si trae porcentajes ('003% Spandex; 097% Cotone'), devolver SOLO el "
    "material con MAYOR porcentaje\n"
    "- Una sola palabra, en espanol: Cotton->Algodon, Polyester->Poliester, "
    "Wool->Lana, Silk->Seda, Viscose->Viscosa, Linen->Lino, Leather->Cuero, "
    "Polyamide->Poliamida, Cashmere->Cachemira, Acrylic->Acrilico"
)


def _traducir(valores, que, reglas, clave=None, forzar=False):
    """
    Traduce una lista de valores. Cachea CADA valor por separado, asi un
    archivo nuevo de la misma marca solo paga los que no vio antes.
    """
    unicos = [v for v in dict.fromkeys(str(x).strip() for x in valores) if v]
    if not unicos:
        return {}

    cache = _cache_leer(f"trad-{que}") or {} if not forzar else {}
    faltan = [v for v in unicos if v.upper() not in cache]
    ULTIMO["del_cache"] += len(unicos) - len(faltan)
    ULTIMO["nuevos"] += len(faltan)

    for i in range(0, len(faltan), LOTE):
        lote = faltan[i:i + LOTE]
        prompt = (
            f"Traduce estos {que} de productos de moda al espanol.\n\n"
            f"REGLAS:\n{reglas}\n\n"
            f"VALORES: {json.dumps(lote, ensure_ascii=False)}\n\n"
            'Devolve un JSON {"ORIGINAL": "traduccion"} con TODOS los valores '
            "de la lista, sin omitir ninguno."
        )
        texto, _ = _pedir(prompt, 60 + 22 * len(lote), clave, SISTEMA)
        ULTIMO["llamadas"] += 1
        for k, v in _json_de(texto).items():
            cache[str(k).strip().upper()] = str(v).strip()

    if faltan:
        _cache_escribir(f"trad-{que}", cache)

    return {v: cache.get(v.upper(), v.capitalize()) for v in unicos}


def traducir_colores(valores, clave=None, forzar=False):
    return _traducir(valores, "colores", REGLAS_COLOR, clave, forzar)


def traducir_materiales(valores, clave=None, forzar=False):
    return _traducir(valores, "materiales", REGLAS_MATERIAL, clave, forzar)


# ===========================================================================
# RESPALDO SIN IA
# ===========================================================================

def material_dominante(texto):
    """
    Saca el material de mayor porcentaje sin llamar a Claude.

    Se usa cuando no hay clave o la llamada falla: es mejor entregar algo
    razonable que cortar todo el proceso.

    >>> material_dominante("003% Spandex; 097% Cotone;")
    'Cotone'
    """
    partes = re.findall(r"0*(\d+(?:[.,]\d+)?)\s*%\s*([^;,%\d]+)", str(texto))
    if not partes:
        return str(texto).strip().capitalize()
    mejor = max(partes, key=lambda p: float(p[0].replace(",", ".")))
    return mejor[1].strip().capitalize()


def estado_cache():
    """Cuantas respuestas hay guardadas, para informar en la ventana."""
    if not CACHE.exists():
        return {"consultas": 0, "colores": 0, "materiales": 0}
    return {
        "consultas": len(list(CACHE.glob("cols-*.json"))),
        "colores": len(_cache_leer("trad-colores") or {}),
        "materiales": len(_cache_leer("trad-materiales") or {}),
    }