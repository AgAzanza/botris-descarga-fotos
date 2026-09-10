#!/usr/bin/env python3
"""
Corre todas las marcas de uno o varios excels, sin tener que acordarse de
que script va con cual.

Lee cada archivo, mira que marcas trae y llama al programa correspondiente.
Si aparece una marca sin programa (por ejemplo FUSTER), lo avisa y sigue.

Se puede usar de dos formas: desde la terminal, o llamando a ejecutar()
desde otro programa (es lo que hace la ventana, app.py).

USO
---
    python runner.py entradas/*.xls
    python runner.py Geox.xls Lanidor.xls --limite 3
    python runner.py entradas/*.xls --salida C:/fotos
"""

import argparse
from pathlib import Path

import comun

# Los modulos de marca se importan explicitamente y no por nombre. Es a
# proposito: asi PyInstaller los detecta al armar el .exe. Con importlib
# no los veria y quedarian afuera del ejecutable.
import geox
import hugo
import lanidor
import mk


# Que modulo atiende cada valor de la columna COD.MARCA del excel.
MODULOS = {
    "GEOX": geox,
    "LANI": lanidor,
    "HUGO": hugo,
    "BOSS": hugo,        # el mismo archivo cubre las dos
    "MK":   mk,
}

# Marcas que un mismo modulo resuelve en una sola corrida.
JUNTAS = {"hugo": ["HUGO", "BOSS"]}


def marcas_de(excel):
    """Que marcas trae el excel y cuantos productos de cada una."""
    try:
        prod = comun.cargar(excel)
    except Exception as e:
        print(f"  no se pudo leer: {e}")
        return {}
    return prod["marca"].value_counts().to_dict()


def ejecutar(excels, salida=None, limite=None, forzar=False):
    """
    Procesa los excels. Devuelve (programas_ok, marcas_sin_programa).

    Llama a cada marca como funcion, no como proceso aparte, para que
    funcione igual dentro del .exe.
    """
    extra = []
    if salida:
        extra += ["--salida", str(salida)]
    if limite:
        extra += ["--limite", str(limite)]
    if forzar:
        extra += ["--forzar"]

    total_ok, sin_programa = 0, []

    for excel in [Path(e) for e in excels]:
        print(f"\n{'=' * 70}\n  {excel.name}\n{'=' * 70}")
        conteo = marcas_de(excel)
        if not conteo:
            continue
        for m, n in conteo.items():
            print(f"  {m:<8} {n:>4} productos"
                  + ("" if m.upper() in MODULOS else "   <-- sin programa"))

        ya = set()
        for marca in conteo:
            modulo = MODULOS.get(marca.upper())
            if modulo is None:
                sin_programa.append(f"{marca} ({excel.name})")
                continue
            nombre = modulo.__name__
            if nombre in ya:
                continue          # hugo ya hizo HUGO y BOSS juntas
            ya.add(nombre)

            cubre = JUNTAS.get(nombre, [marca])
            print(f"\n  --> {nombre}  ({', '.join(cubre)})")
            try:
                modulo.main([str(excel)] + extra)
                total_ok += 1
            except Exception as e:
                print(f"  {nombre} termino con error: {e}")

    print(f"\n{'=' * 70}")
    print(f"  programas ejecutados: {total_ok}")
    if sin_programa:
        print(f"  marcas sin programa : {', '.join(sorted(set(sin_programa)))}")
        print("  (se agregan en MODULOS, arriba de este archivo)")
    return total_ok, sin_programa


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Corre todas las marcas de uno o varios excels.")
    ap.add_argument("excels", nargs="+", type=Path)
    ap.add_argument("--salida", default=None)
    ap.add_argument("--limite", default=None)
    ap.add_argument("--forzar", action="store_true")
    args = ap.parse_args(argv)
    ejecutar(args.excels, args.salida, args.limite, args.forzar)


if __name__ == "__main__":
    main()
