# -*- coding: utf-8 -*-
"""
Fecha: 23/11/2026: Version estable y funcional.

CONSOLIDADOR DE CARTERA (Saldos de Préstamos + Saldos de Descuentos)
=====================================================================
Qué hace, en resumen:
  1. Calcula la fecha de referencia (hoy - 1 día, porque los datos vienen al día anterior).
  2. Entra a la ruta de red:  Año > Mes (nombre) > Día.
  3. Busca los archivos Pr_Saldos_Prestamos_ddmmyyyy y DD_Saldos_Descuentos_ddmmyyyy
     (solo los que cumplen EXACTAMENTE ese formato de nombre).
  4. Toma las columnas requeridas de cada archivo y las apila en una sola tabla.
  5. Agrega las columnas No, Documento, Fecha y Archivo.
  6. Genera un CSV con fecha y hora en el nombre y valida que la cantidad de
     registros del CSV sea igual a la de los archivos originales.

Requisitos: pip install pandas openpyxl   (xlrd solo si hubiera archivos .xls)
Código de salida: 0 = proceso correcto, 1 = error o validación fallida
(útil para el Programador de tareas de Windows).
"""

# ------------------------------------------------------------------ IMPORTS
import logging                      # Para mostrar mensajes en pantalla y guardarlos en un log
import os                           # Para recorrer carpetas y reemplazar/eliminar archivos
import re                           # Expresiones regulares: validar el formato exacto del nombre
import sys                          # Para enviar la salida a pantalla y devolver el código de salida
import unicodedata                  # Para quitar tildes al comparar nombres de carpetas
from datetime import datetime, timedelta   # Manejo de fechas
from pathlib import Path            # Manejo cómodo de rutas de archivos

import pandas as pd                 # Lectura de Excel, consolidación y escritura de CSV

# ------------------------------------------------------------------ CONFIGURACIÓN
# Ruta de red donde están las carpetas por año (ruta UNC completa, con doble barra inicial)

RUTA_ORIGEN = r"C:\Users\HN39280\Desktop\Cartera\CP"
#RUTA_ORIGEN = r"\\172.23.6.81\Listados2021"


# Carpeta donde se guarda el CSV consolidado y el log
RUTA_SALIDA = r"C:\Users\HN39280\Desktop\Cartera\CP"
#RUTA_SALIDA = r"C:\Users\HN39280\Desktop\Cartera\CP"

# True = antes de crear el nuevo CSV se eliminan los CSV anteriores del mismo mes
# (así no se acumulan archivos y se mantiene el esquema "delete / insert").
# Si algún archivo anterior está abierto en Excel, no se puede borrar: solo se avisa.

ELIMINAR_ANTERIORES = True

# Nombres de los meses tal como se llaman las carpetas (sin tildes y en minúscula,
# porque al comparar se normaliza el nombre real de la carpeta)
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

# Columnas que ambos archivos tienen en común (mismo nombre en los dos)
COLS_COMUNES = ["CODIGO_AGENCIA", "CODIGO_SUB_APLICACION", "DESC_MONEDA",
                "CODIGO_ESTADO_CARTERA", "NOMBRE_CLIENTE", "APERTURA",
                "VENCIMIENTO", "TASA_TOTAL", "CAPITAL_TOTAL", "CODIGO_EJECUTIVO",
                "OFICIAL_CREDITO", "AGENCIA_ORIGEN_PRESTAMO", "NOMBRE_AGENCIA"]

# Definición de cada tipo de archivo:
#   regex     -> formato EXACTO del nombre: prefijo + ddmmyyyy (8 dígitos) + extensión.
#                Por eso se descartan nombres como Pr_Saldos_Prestamos_Castigados_03112026
#   columnas  -> columnas que se leen del Excel
#   renombrar -> nombre final que tendrá cada columna en el CSV

TIPOS = {
    "Saldos": {
        "regex": re.compile(r"^Pr_Saldos_Prestamos_(\d{8})\.(xlsx|xlsm|xls)$", re.I),
        "columnas": ["PRESTAMO", "MONTO_INICIAL"] + COLS_COMUNES,
        "renombrar": {"PRESTAMO": "PRESTAMO_NEGOCIACION"},
    },
    "Descuentos": {
        "regex": re.compile(r"^DD_Saldos_Descuentos_(\d{8})\.(xlsx|xlsm|xls)$", re.I),
        "columnas": ["NEGOCIACION", "VALOR_INICIAL"] + COLS_COMUNES,
        # Se unifican los nombres para poder apilar ambos archivos en las mismas columnas
        "renombrar": {"NEGOCIACION": "PRESTAMO_NEGOCIACION", "VALOR_INICIAL": "MONTO_INICIAL"},
    },
}

# Orden final de las columnas del CSV
ORDEN_FINAL = ["No", "Documento", "Fecha", "Archivo", "CODIGO_AGENCIA",
               "CODIGO_SUB_APLICACION", "PRESTAMO_NEGOCIACION", "DESC_MONEDA",
               "CODIGO_ESTADO_CARTERA", "NOMBRE_CLIENTE", "APERTURA", "VENCIMIENTO",
               "TASA_TOTAL", "MONTO_INICIAL", "CAPITAL_TOTAL", "CODIGO_EJECUTIVO",
               "OFICIAL_CREDITO", "AGENCIA_ORIGEN_PRESTAMO", "NOMBRE_AGENCIA"]


# ------------------------------------------------------------------ FUNCIONES AUXILIARES
def norm(texto: str) -> str:
    """Normaliza un texto: quita tildes, espacios sobrantes y lo pasa a minúscula.
    Sirve para que 'Septiembre', 'septiembre ' o 'SEPTIEMBRE' se consideren iguales."""
    t = unicodedata.normalize("NFKD", str(texto))                  # separa letra y tilde
    return "".join(c for c in t if not unicodedata.combining(c)).strip().lower()  # descarta tildes


def buscar_subcarpeta(base: Path, nombres: set):
    """Busca dentro de 'base' una subcarpeta cuyo nombre (normalizado) esté en 'nombres'.
    Devuelve su ruta o None si no existe."""
    for e in os.scandir(base):                                     # recorre el contenido de la carpeta
        if e.is_dir() and norm(e.name) in nombres:                 # solo carpetas con nombre esperado
            return Path(e.path)
    return None


def localizar_archivos(carpeta_mes: Path):
    """Recorre todas las carpetas de día dentro del mes.
    Devuelve:
      - encontrados: {(tipo, fecha): ruta} solo con archivos de formato válido
      - descartados: nombres parecidos que se omitieron por no cumplir el formato"""
    encontrados, descartados = {}, []
    for raiz, _, archivos in os.walk(carpeta_mes):                 # os.walk entra a cada carpeta de día
        for nombre in archivos:
            if nombre.startswith("~$"):                            # archivos temporales de Excel abiertos
                continue
            coincide = False                                       # ¿el nombre cumple algún formato válido?
            for tipo, cfg in TIPOS.items():
                m = cfg["regex"].match(nombre)                     # compara con prefijo_ddmmyyyy.ext
                if not m:
                    continue
                try:
                    # Convierte los 8 dígitos a fecha real; falla si es una fecha imposible (p. ej. 45132026)
                    fecha = datetime.strptime(m.group(1), "%d%m%Y").date()
                except ValueError:
                    continue
                coincide = True
                ruta = Path(raiz) / nombre
                clave = (tipo, fecha)
                if clave in encontrados:                           # mismo archivo y fecha en dos carpetas
                    logging.warning("Archivo duplicado para %s %s; se usa el más reciente.", tipo, fecha)
                    if ruta.stat().st_mtime < encontrados[clave].stat().st_mtime:
                        continue                                   # se conserva el modificado más recientemente
                encontrados[clave] = ruta
            # Si empieza con el prefijo pero no cumple el formato, se registra como omitido
            if not coincide and nombre.lower().startswith(("pr_saldos_prestamos_", "dd_saldos_descuentos_")):
                descartados.append(nombre)
    return encontrados, descartados


def leer(ruta: Path, tipo: str) -> pd.DataFrame:
    """Lee un Excel y devuelve solo las columnas requeridas, ya con los nombres finales."""
    cfg = TIPOS[tipo]
    requeridas = {c.upper() for c in cfg["columnas"]}              # columnas esperadas, en mayúscula
    df = pd.read_excel(
        ruta,
        # Lee únicamente las columnas pedidas (ignora el resto: más rápido y liviano)
        usecols=lambda c: str(c).strip().upper() in requeridas,
        # El número de préstamo/negociación se lee como texto para no perder ceros ni formato
        dtype={"PRESTAMO": str, "NEGOCIACION": str},
    )
    df.columns = [str(c).strip().upper() for c in df.columns]      # limpia espacios en los encabezados
    faltan = requeridas - set(df.columns)                          # ¿falta alguna columna requerida?
    if faltan:
        raise ValueError(f"{ruta.name}: faltan columnas {sorted(faltan)}")
    # Deja las columnas en el orden definido y aplica los nombres finales
    return df[cfg["columnas"]].rename(columns=cfg["renombrar"])


def total_capital(serie: pd.Series) -> float:
    """Suma de CAPITAL_TOTAL (valores no numéricos cuentan como 0).
    Se usa como control adicional además del conteo de registros."""
    return float(pd.to_numeric(serie, errors="coerce").fillna(0).sum())


# ------------------------------------------------------------------ PROCESO PRINCIPAL
def main() -> int:
    # ---- Preparación: carpeta de salida y log
    salida = Path(RUTA_SALIDA)
    salida.mkdir(parents=True, exist_ok=True)                      # crea la carpeta si no existe
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",            # cada línea lleva fecha/hora y nivel
        handlers=[logging.FileHandler(salida / "consolidar_cartera.log", encoding="utf-8"),  # al archivo log
                  logging.StreamHandler(sys.stdout)],              # y a la pantalla
    )
    log = logging.info                                             # atajo para escribir mensajes informativos

    # ---- Fecha de referencia y nombre del archivo de salida
    ahora = datetime.now()
    ref = (ahora - timedelta(days=1)).date()                       # datos al día anterior
    
    #prefijo = f"Consolidado_Saldos_Descuentos_{ref.year}{ref.month:02d}"              # ej.: Cartera_CP_202611
    prefijo = f"Consolidado_Saldos_Descuentos_"
    # El nombre lleva fecha y hora de ejecución: Cartera_CP_202611_20261123_153045.csv
    destino = salida / f"{prefijo}_{ahora:%Y%m%d_%H%M%S}.csv"

    log("=== INICIO DEL PROCESO === Fecha de referencia: %s", ref)
    # ---- PASO 1: eliminar CSV anteriores del mismo mes (delete)
    log("[1/5] Eliminando archivos generados anteriormente...")
    if ELIMINAR_ANTERIORES:
        previos = list(salida.glob(f"{prefijo}*.csv"))             # CSV previos del mes en proceso
        for p in previos:
            try:
                p.unlink()                                         # elimina el archivo
                log("      Eliminado: %s", p.name)
            except PermissionError:                                # abierto en Excel u otro programa
                logging.warning("      No se pudo eliminar %s (¿está abierto?). Se continúa.", p.name)
        if not previos:
            log("      No existían archivos previos.")

    # ---- PASO 2: ubicar carpetas y archivos de origen
    log("[2/5] Buscando carpetas y archivos de origen...")
    carpeta_anio = buscar_subcarpeta(Path(RUTA_ORIGEN), {str(ref.year)})   # carpeta del año, ej. "2026"
    if not carpeta_anio:
        logging.error("No existe la carpeta del año %s en %s", ref.year, RUTA_ORIGEN)
        return 1
    # Nombre de la carpeta del mes (acepta también "setiembre" por si se escribe así)
    nombres_mes = {MESES[ref.month - 1]} | ({"setiembre"} if ref.month == 9 else set())
    carpeta_mes = buscar_subcarpeta(carpeta_anio, nombres_mes)
    if not carpeta_mes:
        logging.error("No existe la carpeta del mes %s en %s", MESES[ref.month - 1], carpeta_anio)
        return 1
    log("      Carpeta de mes: %s", carpeta_mes)

    todos, descartados = localizar_archivos(carpeta_mes)           # busca en todas las carpetas de día
    for n in sorted(descartados):
        log("      Omitido por formato de nombre: %s", n)
    # Solo se consolida UN mes: se conservan los archivos del mismo año/mes de la fecha de referencia
    # y con fecha menor o igual a ella (nunca se mezclan meses)
    archivos = {k: v for k, v in todos.items()
                if (k[1].year, k[1].month) == (ref.year, ref.month) and k[1] <= ref}
    for tipo, fecha in sorted(set(todos) - set(archivos)):
        logging.warning("      Fuera del mes en proceso, se omite: %s %s", tipo, fecha)
    if not archivos:
        logging.error("No se encontraron archivos válidos en %s", carpeta_mes)
        return 1

    fechas = sorted({f for _, f in archivos})                      # fechas distintas encontradas
    log("      Archivos válidos: %d | Último día cargado: %s (desde %s)",
        len(archivos), fechas[-1], fechas[0])                      # fechas[-1] = último archivo del mes

    # Informativo: días hábiles (lun-vie) sin archivo. Fines de semana no se reportan.
    d, faltantes = fechas[0].replace(day=1), []
    while d <= fechas[-1]:
        if d.weekday() < 5 and d not in fechas:                    # weekday < 5 = lunes a viernes
            faltantes.append(d.strftime("%d/%m"))
        d += timedelta(days=1)
    if faltantes:
        log("      Días hábiles sin archivo (feriados u omisiones): %s", ", ".join(faltantes))
    # Aviso si una fecha tiene un tipo de archivo pero le falta el otro
    for f in fechas:
        for tipo in TIPOS:
            if (tipo, f) not in archivos:
                logging.warning("      Falta archivo de %s para la fecha %s", tipo, f)

    # ---- PASO 3: leer cada archivo
    log("[3/5] Leyendo archivos...")
    claves = sorted(archivos, key=lambda k: (k[1], k[0]))          # orden por fecha y luego por tipo
    partes, control = [], {}                                       # partes = tablas leídas; control = conteos originales
    for i, (tipo, fecha) in enumerate(claves, 1):
        ruta = archivos[(tipo, fecha)]
        log("      (%d/%d) Leyendo %s ...", i, len(claves), ruta.name)   # avance archivo por archivo
        df = leer(ruta, tipo)
        # Se guarda el conteo y la suma de capital del ORIGINAL para validar después
        control[ruta.name] = {"filas": len(df), "capital": total_capital(df["CAPITAL_TOTAL"])}
        df.insert(0, "Archivo", ruta.name)                         # nombre del archivo de origen
        df.insert(0, "Fecha", pd.Timestamp(fecha))                 # fecha tomada del nombre del archivo
        df.insert(0, "Documento", tipo)                            # "Saldos" o "Descuentos"
        partes.append(df)
        log("            %d registros", len(df))

    # ---- PASO 4: consolidar y generar el CSV (insert)
    log("[4/5] Consolidando datos y generando CSV (puede tardar)...")
    final = pd.concat(partes, ignore_index=True)                   # apila todas las tablas en una sola
    del partes                                                     # libera memoria
    final.insert(0, "No", range(1, len(final) + 1))                # correlativo de línea desde 1
    final = final[ORDEN_FINAL]                                     # ordena las columnas
    # utf-8-sig: Excel reconoce bien tildes y ñ; date_format: fechas como dd/mm/yyyy
    final.to_csv(destino, index=False, encoding="utf-8-sig", date_format="%d/%m/%Y")
    total_generado = len(final)                                    # filas que se escribieron
    del final
    log("      CSV generado: %s", destino)

    # ---- PASO 5: validar contra el archivo ya escrito en disco
    log("[5/5] Validando registros entre archivos originales y CSV generado...")
    conteo, capital = {}, {}                                       # lo que realmente quedó en el CSV
    # Se relee el CSV por bloques (chunks) para no cargar todo en memoria
    for chunk in pd.read_csv(destino, usecols=["Archivo", "CAPITAL_TOTAL"],
                             chunksize=500_000, encoding="utf-8-sig"):
        for nombre, g in chunk.groupby("Archivo"):
            conteo[nombre] = conteo.get(nombre, 0) + len(g)
            capital[nombre] = capital.get(nombre, 0.0) + total_capital(g["CAPITAL_TOTAL"])

    ok = True
    for nombre, c in control.items():                              # compara archivo por archivo
        if conteo.get(nombre, 0) != c["filas"] or abs(capital.get(nombre, 0.0) - c["capital"]) > 0.01:
            ok = False
            logging.error("DIFERENCIA en %s: origen=%d filas | CSV=%d filas",
                          nombre, c["filas"], conteo.get(nombre, 0))
    total_origen = sum(c["filas"] for c in control.values())       # total de registros en los originales
    total_csv = sum(conteo.values())                               # total de registros en el CSV
    if not (total_origen == total_generado == total_csv):          # los tres conteos deben coincidir
        ok = False

    print()
    log("Archivos procesados       : %d", len(control))
    log("Total registros originales: %s", f"{total_origen:,}")
    log("Total registros en CSV    : %s", f"{total_csv:,}")

    if not ok:
        logging.error("PROCESO TERMINADO CON ERRORES: la cantidad de registros NO coincide.")
        return 1

    log("Validación correcta: la cantidad de registros coincide.")
    log("=== PROCESO CONCLUIDO CON EXITO === %s", destino)
    return 0


# ------------------------------------------------------------------ PUNTO DE ENTRADA
if __name__ == "__main__":
    try:
        sys.exit(main())                                           # ejecuta y devuelve 0 (OK) o 1 (error)
    except Exception:
        # Cualquier error inesperado (red caída, columna faltante, etc.) queda en pantalla y en el log
        logging.exception("PROCESO TERMINADO CON ERRORES (error no controlado)")
        sys.exit(1)
