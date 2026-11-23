# -*- coding: utf-8 -*-
"""
Consolida Pr_Saldos_Prestamos_ddmmyyyy y DD_Saldos_Descuentos_ddmmyyyy
del mes en curso (fecha de referencia = hoy - 1 día) en un solo CSV.

Requisitos: pip install pandas openpyxl   (xlrd si hubiera archivos .xls)
Código de salida: 0 = OK, 1 = error/validación fallida.
"""
import logging
import os
import re
import sys
import unicodedata
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

# ------------------------------------------------------------------ CONFIG
RUTA_ORIGEN = r"\\172.23.6.81\Listados2021"
RUTA_SALIDA = r"C:\Users\HN39280\Desktop\Cartera\CP"

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

COLS_COMUNES = ["CODIGO_AGENCIA", "CODIGO_SUB_APLICACION", "DESC_MONEDA",
                "CODIGO_ESTADO_CARTERA", "NOMBRE_CLIENTE", "APERTURA",
                "VENCIMIENTO", "TASA_TOTAL", "CAPITAL_TOTAL", "CODIGO_EJECUTIVO",
                "OFICIAL_CREDITO", "AGENCIA_ORIGEN_PRESTAMO", "NOMBRE_AGENCIA"]

# El nombre debe ser EXACTAMENTE <prefijo>_ddmmyyyy.<ext>
# (Pr_Saldos_Prestamos_Castigados_03112026, Pr_Saldos_Prestamos_Humberto_xxx, etc. se descartan)
TIPOS = {
    "Saldos": {
        "regex": re.compile(r"^Pr_Saldos_Prestamos_(\d{8})\.(xlsx|xlsm|xls)$", re.I),
        "columnas": ["PRESTAMO", "MONTO_INICIAL"] + COLS_COMUNES,
        "renombrar": {"PRESTAMO": "PRESTAMO_NEGOCIACION"},
    },
    "Descuentos": {
        "regex": re.compile(r"^DD_Saldos_Descuentos_(\d{8})\.(xlsx|xlsm|xls)$", re.I),
        "columnas": ["NEGOCIACION", "VALOR_INICIAL"] + COLS_COMUNES,
        "renombrar": {"NEGOCIACION": "PRESTAMO_NEGOCIACION", "VALOR_INICIAL": "MONTO_INICIAL"},
    },
}

ORDEN_FINAL = ["No", "Documento", "Fecha", "Archivo", "CODIGO_AGENCIA",
               "CODIGO_SUB_APLICACION", "PRESTAMO_NEGOCIACION", "DESC_MONEDA",
               "CODIGO_ESTADO_CARTERA", "NOMBRE_CLIENTE", "APERTURA", "VENCIMIENTO",
               "TASA_TOTAL", "MONTO_INICIAL", "CAPITAL_TOTAL", "CODIGO_EJECUTIVO",
               "OFICIAL_CREDITO", "AGENCIA_ORIGEN_PRESTAMO", "NOMBRE_AGENCIA"]


# ------------------------------------------------------------------ UTILES
def norm(texto: str) -> str:
    t = unicodedata.normalize("NFKD", str(texto))
    return "".join(c for c in t if not unicodedata.combining(c)).strip().lower()


def buscar_subcarpeta(base: Path, nombres: set):
    for e in os.scandir(base):
        if e.is_dir() and norm(e.name) in nombres:
            return Path(e.path)
    return None


def localizar_archivos(carpeta_mes: Path):
    """Devuelve ({(tipo, fecha): ruta}, [nombres descartados por formato])."""
    encontrados, descartados = {}, []
    for raiz, _, archivos in os.walk(carpeta_mes):
        for nombre in archivos:
            if nombre.startswith("~$"):
                continue
            coincide = False
            for tipo, cfg in TIPOS.items():
                m = cfg["regex"].match(nombre)
                if not m:
                    continue
                try:
                    fecha = datetime.strptime(m.group(1), "%d%m%Y").date()
                except ValueError:
                    continue
                coincide = True
                ruta = Path(raiz) / nombre
                clave = (tipo, fecha)
                if clave in encontrados:
                    logging.warning("Archivo duplicado para %s %s; se usa el más reciente.", tipo, fecha)
                    if ruta.stat().st_mtime < encontrados[clave].stat().st_mtime:
                        continue
                encontrados[clave] = ruta
            if not coincide and nombre.lower().startswith(("pr_saldos_prestamos_", "dd_saldos_descuentos_")):
                descartados.append(nombre)
    return encontrados, descartados


def leer(ruta: Path, tipo: str) -> pd.DataFrame:
    cfg = TIPOS[tipo]
    requeridas = {c.upper() for c in cfg["columnas"]}
    df = pd.read_excel(
        ruta,
        usecols=lambda c: str(c).strip().upper() in requeridas,
        dtype={"PRESTAMO": str, "NEGOCIACION": str},
    )
    df.columns = [str(c).strip().upper() for c in df.columns]
    faltan = requeridas - set(df.columns)
    if faltan:
        raise ValueError(f"{ruta.name}: faltan columnas {sorted(faltan)}")
    return df[cfg["columnas"]].rename(columns=cfg["renombrar"])


def total_capital(serie: pd.Series) -> float:
    return float(pd.to_numeric(serie, errors="coerce").fillna(0).sum())


# ------------------------------------------------------------------ PROCESO
def main() -> int:
    salida = Path(RUTA_SALIDA)
    salida.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[logging.FileHandler(salida / "consolidar_cartera.log", encoding="utf-8"),
                  logging.StreamHandler(sys.stdout)],
    )
    log = logging.info

    ref = (datetime.now() - timedelta(days=1)).date()  # los datos están al día anterior
    destino = salida / f"Cartera_CP_{ref.year}{ref.month:02d}.csv"

    log("=== INICIO DEL PROCESO === Fecha de referencia: %s", ref)

    # Delete: se elimina el archivo generado anteriormente
    log("[1/5] Eliminando archivo generado anteriormente...")
    try:
        if destino.exists():
            destino.unlink()
            log("      Eliminado: %s", destino)
        else:
            log("      No existía archivo previo.")
    except PermissionError:
        logging.error("No se pudo eliminar %s. ¿Está abierto en Excel? Ciérrelo y reintente.", destino)
        return 1

    log("[2/5] Buscando carpetas y archivos de origen...")
    carpeta_anio = buscar_subcarpeta(Path(RUTA_ORIGEN), {str(ref.year)})
    if not carpeta_anio:
        logging.error("No existe la carpeta del año %s en %s", ref.year, RUTA_ORIGEN)
        return 1
    nombres_mes = {MESES[ref.month - 1]} | ({"setiembre"} if ref.month == 9 else set())
    carpeta_mes = buscar_subcarpeta(carpeta_anio, nombres_mes)
    if not carpeta_mes:
        logging.error("No existe la carpeta del mes %s en %s", MESES[ref.month - 1], carpeta_anio)
        return 1
    log("      Carpeta de mes: %s", carpeta_mes)

    todos, descartados = localizar_archivos(carpeta_mes)
    for n in sorted(descartados):
        log("      Omitido por formato de nombre: %s", n)
    archivos = {k: v for k, v in todos.items()
                if (k[1].year, k[1].month) == (ref.year, ref.month) and k[1] <= ref}
    for tipo, fecha in sorted(set(todos) - set(archivos)):
        logging.warning("      Fuera del mes en proceso, se omite: %s %s", tipo, fecha)
    if not archivos:
        logging.error("No se encontraron archivos válidos en %s", carpeta_mes)
        return 1

    fechas = sorted({f for _, f in archivos})
    log("      Archivos válidos: %d | Último día cargado: %s (desde %s)",
        len(archivos), fechas[-1], fechas[0])

    d, faltantes = fechas[0].replace(day=1), []
    while d <= fechas[-1]:
        if d.weekday() < 5 and d not in fechas:
            faltantes.append(d.strftime("%d/%m"))
        d += timedelta(days=1)
    if faltantes:
        log("      Días hábiles sin archivo (feriados u omisiones): %s", ", ".join(faltantes))
    for f in fechas:
        for tipo in TIPOS:
            if (tipo, f) not in archivos:
                logging.warning("      Falta archivo de %s para la fecha %s", tipo, f)

    # Lectura
    log("[3/5] Leyendo archivos...")
    claves = sorted(archivos, key=lambda k: (k[1], k[0]))
    partes, control = [], {}
    for i, (tipo, fecha) in enumerate(claves, 1):
        ruta = archivos[(tipo, fecha)]
        log("      (%d/%d) Leyendo %s ...", i, len(claves), ruta.name)
        df = leer(ruta, tipo)
        control[ruta.name] = {"filas": len(df), "capital": total_capital(df["CAPITAL_TOTAL"])}
        df.insert(0, "Archivo", ruta.name)
        df.insert(0, "Fecha", pd.Timestamp(fecha))
        df.insert(0, "Documento", tipo)
        partes.append(df)
        log("            %d registros", len(df))

    # Consolidación + Insert
    log("[4/5] Consolidando datos y generando CSV (puede tardar)...")
    final = pd.concat(partes, ignore_index=True)
    del partes
    final.insert(0, "No", range(1, len(final) + 1))
    final = final[ORDEN_FINAL]
    final.to_csv(destino, index=False, encoding="utf-8-sig", date_format="%d/%m/%Y")
    total_generado = len(final)
    del final
    log("      CSV generado: %s", destino)

    # Validación contra el archivo ya escrito en disco
    log("[5/5] Validando registros entre archivos originales y CSV generado...")
    conteo, capital = {}, {}
    for chunk in pd.read_csv(destino, usecols=["Archivo", "CAPITAL_TOTAL"],
                             chunksize=500_000, encoding="utf-8-sig"):
        for nombre, g in chunk.groupby("Archivo"):
            conteo[nombre] = conteo.get(nombre, 0) + len(g)
            capital[nombre] = capital.get(nombre, 0.0) + total_capital(g["CAPITAL_TOTAL"])

    ok = True
    for nombre, c in control.items():
        if conteo.get(nombre, 0) != c["filas"] or abs(capital.get(nombre, 0.0) - c["capital"]) > 0.01:
            ok = False
            logging.error("DIFERENCIA en %s: origen=%d filas | CSV=%d filas",
                          nombre, c["filas"], conteo.get(nombre, 0))
    total_origen = sum(c["filas"] for c in control.values())
    total_csv = sum(conteo.values())
    if not (total_origen == total_generado == total_csv):
        ok = False

    print()
    log("Archivos procesados      : %d", len(control))
    log("Total registros originales: %s", f"{total_origen:,}")
    log("Total registros en CSV     : %s", f"{total_csv:,}")

    if not ok:
        logging.error("PROCESO TERMINADO CON ERRORES: la cantidad de registros NO coincide.")
        return 1

    log("Validación correcta: la cantidad de registros coincide.")
    log("=== PROCESO CONCLUIDO CON EXITO === %s", destino)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        logging.exception("PROCESO TERMINADO CON ERRORES (error no controlado)")
        sys.exit(1)
