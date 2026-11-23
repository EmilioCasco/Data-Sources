# -*- coding: utf-8 -*-
"""
Consolida Pr_Saldos_Prestamos_ddmmyyyy y DD_Saldos_Descuentos_ddmmyyyy
del mes en curso (fecha de referencia = hoy - 1 día) en un solo Excel.

Requisitos: pip install pandas openpyxl   (xlrd si hubiera archivos .xls)
Código de salida: 0 = OK, 1 = error/validación fallida (útil para el programador de tareas).
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

TIPOS = {
    "Saldos": {
        "regex": re.compile(r"^Pr_Saldos_Prestamos_(\d{8})\.(xlsx|xlsm|xls)$", re.I),
        "columnas": ["PRESTAMO", "MONTO_INICIAL"] + COLS_COMUNES,
        "renombrar": {},
    },
    "Descuentos": {
        "regex": re.compile(r"^DD_Saldos_Descuentos_(\d{8})\.(xlsx|xlsm|xls)$", re.I),
        "columnas": ["NEGOCIACION", "VALOR_INICIAL"] + COLS_COMUNES,
        # Se unifican para poder apilar ambos archivos en las mismas columnas
        "renombrar": {"NEGOCIACION": "PRESTAMO", "VALOR_INICIAL": "MONTO_INICIAL"},
    },
}

ORDEN_FINAL = (["Documento", "Fecha", "Archivo", "CODIGO_AGENCIA", "CODIGO_SUB_APLICACION",
                "PRESTAMO", "DESC_MONEDA", "CODIGO_ESTADO_CARTERA", "NOMBRE_CLIENTE",
                "APERTURA", "VENCIMIENTO", "TASA_TOTAL", "MONTO_INICIAL", "CAPITAL_TOTAL",
                "CODIGO_EJECUTIVO", "OFICIAL_CREDITO", "AGENCIA_ORIGEN_PRESTAMO",
                "NOMBRE_AGENCIA"])

MAX_FILAS_EXCEL = 1_048_575


# ------------------------------------------------------------------ UTILES
def norm(texto: str) -> str:
    t = unicodedata.normalize("NFKD", str(texto))
    return "".join(c for c in t if not unicodedata.combining(c)).strip().lower()


def buscar_subcarpeta(base: Path, nombres: set) -> Path | None:
    for e in os.scandir(base):
        if e.is_dir() and norm(e.name) in nombres:
            return Path(e.path)
    return None


def localizar_archivos(carpeta_mes: Path) -> dict:
    """Recorre las carpetas de día y devuelve {(tipo, fecha): ruta}.
    Solo acepta nombres con formato exacto <prefijo>_ddmmyyyy.<ext> y fecha válida."""
    encontrados = {}
    for raiz, _, archivos in os.walk(carpeta_mes):
        for nombre in archivos:
            if nombre.startswith("~$"):
                continue
            for tipo, cfg in TIPOS.items():
                m = cfg["regex"].match(nombre)
                if not m:
                    continue
                try:
                    fecha = datetime.strptime(m.group(1), "%d%m%Y").date()
                except ValueError:
                    logging.warning("Fecha inválida en nombre, se omite: %s", nombre)
                    continue
                ruta = Path(raiz) / nombre
                clave = (tipo, fecha)
                if clave in encontrados:
                    logging.warning("Archivo duplicado para %s %s; se usa el más reciente.", tipo, fecha)
                    if ruta.stat().st_mtime < encontrados[clave].stat().st_mtime:
                        continue
                encontrados[clave] = ruta
    return encontrados


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


def total_capital(df: pd.DataFrame) -> float:
    return float(pd.to_numeric(df["CAPITAL_TOTAL"], errors="coerce").fillna(0).sum())


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

    ref = (datetime.now() - timedelta(days=1)).date()  # datos al día anterior
    logging.info("Fecha de referencia: %s", ref)

    # Año -> Mes (nombre) de la fecha de referencia
    carpeta_anio = buscar_subcarpeta(Path(RUTA_ORIGEN), {str(ref.year)})
    if not carpeta_anio:
        logging.error("No existe la carpeta del año %s en %s", ref.year, RUTA_ORIGEN)
        return 1
    nombres_mes = {MESES[ref.month - 1]}
    if ref.month == 9:
        nombres_mes.add("setiembre")
    carpeta_mes = buscar_subcarpeta(carpeta_anio, nombres_mes)
    if not carpeta_mes:
        logging.error("No existe la carpeta del mes %s en %s", MESES[ref.month - 1], carpeta_anio)
        return 1
    logging.info("Carpeta de mes: %s", carpeta_mes)

    # Archivos válidos, solo del mes/año de referencia y hasta la fecha de referencia
    todos = localizar_archivos(carpeta_mes)
    archivos = {k: v for k, v in todos.items()
                if (k[1].year, k[1].month) == (ref.year, ref.month) and k[1] <= ref}
    for (tipo, fecha) in sorted(set(todos) - set(archivos)):
        logging.warning("Fuera del mes en proceso, se omite: %s %s", tipo, fecha)
    if not archivos:
        logging.error("No se encontraron archivos válidos en %s", carpeta_mes)
        return 1

    fechas = sorted({f for _, f in archivos})
    logging.info("Último archivo cargado: %s (rango %s a %s)", fechas[-1], fechas[0], fechas[-1])

    # Informativo: días hábiles sin archivo (fines de semana no se generan)
    d, faltantes = fechas[0].replace(day=1), []
    while d <= fechas[-1]:
        if d.weekday() < 5 and d not in fechas:
            faltantes.append(d.strftime("%d/%m"))
        d += timedelta(days=1)
    if faltantes:
        logging.info("Días hábiles sin archivo (feriados u omisiones): %s", ", ".join(faltantes))
    for f in fechas:
        for tipo in TIPOS:
            if (tipo, f) not in archivos:
                logging.warning("Falta archivo de %s para la fecha %s", tipo, f)

    # Lectura y fusión (se reconstruye todo desde el día 1 en cada corrida)
    partes, control = [], []
    for (tipo, fecha) in sorted(archivos, key=lambda k: (k[1], k[0])):
        ruta = archivos[(tipo, fecha)]
        df = leer(ruta, tipo)
        control.append({"Archivo": ruta.name, "filas": len(df), "capital": total_capital(df)})
        df.insert(0, "Archivo", ruta.name)
        df.insert(0, "Fecha", pd.Timestamp(fecha))
        df.insert(0, "Documento", tipo)
        partes.append(df)
        logging.info("Leído %-40s %10d filas", ruta.name, len(df))

    final = pd.concat(partes, ignore_index=True)[ORDEN_FINAL]

    # ---- Validaciones de registros
    ok = True
    por_archivo = final.groupby("Archivo").agg(filas=("Archivo", "size"),
                                               capital=("CAPITAL_TOTAL", lambda s: total_capital(
                                                   pd.DataFrame({"CAPITAL_TOTAL": s}))))
    for c in control:
        nuevo = por_archivo.loc[c["Archivo"]]
        if nuevo["filas"] != c["filas"] or abs(nuevo["capital"] - c["capital"]) > 0.01:
            ok = False
            logging.error("DIFERENCIA en %s: origen=%s filas/%.2f, consolidado=%s filas/%.2f",
                          c["Archivo"], c["filas"], c["capital"], nuevo["filas"], nuevo["capital"])
    total_origen = sum(c["filas"] for c in control)
    if len(final) != total_origen:
        ok = False
        logging.error("Total filas: origen=%d, consolidado=%d", total_origen, len(final))
    if len(final) > MAX_FILAS_EXCEL:
        logging.error("El consolidado (%d filas) excede el límite de Excel; use CSV.", len(final))
        return 1
    if not ok:
        logging.error("Validación fallida: no se genera el archivo de salida.")
        return 1

    # ---- Escritura (reemplaza el consolidado anterior del mes)
    destino = salida / f"Cartera_CP_{ref.year}{ref.month:02d}.xlsx"
    temporal = destino.with_suffix(".tmp.xlsx")
    with pd.ExcelWriter(temporal, engine="openpyxl", datetime_format="dd/mm/yyyy") as w:
        final.to_excel(w, index=False, sheet_name="Consolidado")
    os.replace(temporal, destino)

    logging.info("OK: %d filas (%d archivos) -> %s", len(final), len(control), destino)
    logging.info("Validación de registros: origen = consolidado (%d filas)", total_origen)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        logging.exception("Error no controlado")
        sys.exit(1)
