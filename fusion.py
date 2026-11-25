# -*- coding: utf-8 -*-
"""
Fecha: 25/11/2026
CONSOLIDADOR DE CARTERA (Saldos de Préstamos + Saldos de Descuentos)
=====================================================================
Qué hace, en resumen:
1. Calcula la fecha de referencia (hoy - 1 día, porque los datos vienen al día anterior).
2. Entra a la ruta de red: Año > Mes (nombre) > Día.
3. Busca los archivos Pr_Saldos_Prestamos_ddmmyyyy y DD_Saldos_Descuentos_ddmmyyyy
   (solo los que cumplen EXACTAMENTE ese formato de nombre).
4. VERIFICA que ambos archivos existan y estén actualizados a la MISMA fecha.
   Si no es así, aborta el proceso con un mensaje de error (no se genera nada).
5. Toma las columnas requeridas de cada archivo y las apila en una sola tabla.
6. Agrega las columnas No, Documento, Fecha y Archivo.
7. Genera el consolidado en PARTES de máximo FILAS_POR_PARTE registros cada una
   (para que Excel pueda abrirlas, ya que una hoja admite 1,048,576 filas):
   ConsolidadoMensual_Prestamos_yyyymmdd_parte01.csv, _parte02.csv, ...
   (yyyymmdd = fecha hasta la que están actualizados los datos). Valida que la suma
   de registros de todas las partes sea igual a la de los archivos originales.
8. Los mensajes del proceso se muestran en pantalla y se guardan en Consolidado_Log.log
   (en la carpeta de salida).

Requisitos: pip install pandas openpyxl python-calamine
(python-calamine se usa para leer los archivos .xls; reemplaza a xlrd)

Código de salida: 0 = proceso correcto, 1 = error o validación fallida
(útil para el Programador de tareas de Windows).
"""

# ------------------------------------------------------------------ IMPORTS
import logging          # Para mostrar mensajes en pantalla y armar el log de la ejecución
import os               # Para recorrer carpetas
import re               # Expresiones regulares: validar el formato exacto del nombre
import sys              # Para enviar la salida a pantalla y devolver el código de salida
import unicodedata      # Para quitar tildes al comparar nombres de carpetas
from datetime import datetime, timedelta   # Manejo de fechas
from pathlib import Path                   # Manejo cómodo de rutas de archivos

import pandas as pd     # Lectura de Excel, consolidación y escritura de CSV

# ------------------------------------------------------------------ CONFIGURACIÓN
# Ruta de red donde están las carpetas por año (ruta UNC completa, con doble barra inicial)
#RUTA_ORIGEN = r"C:\Users\HN39280\Desktop\Cartera\CP"
RUTA_ORIGEN = r"\\172.23.6.81\Listados2021"

# Carpeta donde se guardan los CSV consolidados y el log
RUTA_SALIDA = r"C:\Users\HN39280\Desktop\Cartera\CP"

# Prefijo del archivo de salida: <PREFIJO_SALIDA>_yyyymmdd_parteNN.csv
PREFIJO_SALIDA = "ConsolidadoMensual_Prestamos"

# Máximo de registros (sin contar el encabezado) por cada CSV de salida.
# Debe ser menor a 1,048,575 (límite de filas de datos de una hoja de Excel).
FILAS_POR_PARTE = 1_000_000

# Archivo de log (se guarda en la carpeta de salida)
NOMBRE_LOG = "Consolidado_Log.log"

# True = al generar los nuevos CSV también se eliminan los consolidados de OTRAS fechas del
# mismo mes (así solo quedan los más recientes). Los de la misma fecha SIEMPRE se eliminan y se
# vuelven a crear (incluidas partes sobrantes de una corrida anterior). Si algún archivo
# anterior está abierto en Excel, no se puede borrar: solo se avisa.
ELIMINAR_ANTERIORES = False

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
#   nombre    -> prefijo del archivo (solo se usa en los mensajes)
#   regex     -> formato EXACTO del nombre: prefijo + ddmmyyyy (8 dígitos) + extensión.
#                Por eso se descartan nombres como Pr_Saldos_Prestamos_Castigados_03112026
#   columnas  -> columnas que se leen del Excel
#   renombrar -> nombre final que tendrá cada columna en el CSV
TIPOS = {
    "Saldos": {
        "nombre": "Pr_Saldos_Prestamos",
        "regex": re.compile(r"^Pr_Saldos_Prestamos_(\d{8})\.(xlsx|xlsm|xls)$", re.I),
        "columnas": ["PRESTAMO", "MONTO_INICIAL"] + COLS_COMUNES,
        "renombrar": {"PRESTAMO": "PRESTAMO_NEGOCIACION"},
    },
    "Descuentos": {
        "nombre": "DD_Saldos_Descuentos",
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
    t = unicodedata.normalize("NFKD", str(texto))                              # separa letra y tilde
    return "".join(c for c in t if not unicodedata.combining(c)).strip().lower()  # descarta tildes


def buscar_subcarpeta(base: Path, nombres: set):
    """Busca dentro de 'base' una subcarpeta cuyo nombre (normalizado) esté en 'nombres'.
    Devuelve su ruta o None si no existe."""
    for e in os.scandir(base):                              # recorre el contenido de la carpeta
        if e.is_dir() and norm(e.name) in nombres:          # solo carpetas con nombre esperado
            return Path(e.path)
    return None


def localizar_archivos(carpeta_mes: Path):
    """Recorre todas las carpetas de día dentro del mes.
    Devuelve:
      - encontrados: {(tipo, fecha): ruta} solo con archivos de formato válido
      - descartados: nombres parecidos que se omitieron por no cumplir el formato"""
    encontrados, descartados = {}, []
    for raiz, _, archivos in os.walk(carpeta_mes):          # os.walk entra a cada carpeta de día
        for nombre in archivos:
            if nombre.startswith("~$"):                     # archivos temporales de Excel abiertos
                continue
            coincide = False                                # ¿el nombre cumple algún formato válido?
            for tipo, cfg in TIPOS.items():
                m = cfg["regex"].match(nombre)              # compara con prefijo_ddmmyyyy.ext
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
                if clave in encontrados:                    # mismo archivo y fecha en dos carpetas
                    logging.warning("Archivo duplicado para %s %s; se usa el más reciente.", tipo, fecha)
                    if ruta.stat().st_mtime < encontrados[clave].stat().st_mtime:
                        continue                            # se conserva el modificado más recientemente
                encontrados[clave] = ruta
            # Si empieza con el prefijo pero no cumple el formato, se registra como omitido
            if not coincide and nombre.lower().startswith(("pr_saldos_prestamos_", "dd_saldos_descuentos_")):
                descartados.append(nombre)
    return encontrados, descartados


def leer(ruta: Path, tipo: str) -> pd.DataFrame:
    """Lee un Excel y devuelve solo las columnas requeridas, ya con los nombres finales."""
    cfg = TIPOS[tipo]
    requeridas = {c.upper() for c in cfg["columnas"]}       # columnas esperadas, en mayúscula
    # Los .xls se leen con 'calamine' (xlrd está obsoleto y genera advertencias);
    # los .xlsx/.xlsm siguen con el motor por defecto (openpyxl)
    motor = "calamine" if ruta.suffix.lower() == ".xls" else None
    df = pd.read_excel(
        ruta,
        engine=motor,
        # Lee únicamente las columnas pedidas (ignora el resto: más rápido y liviano)
        usecols=lambda c: str(c).strip().upper() in requeridas,
        # El número de préstamo/negociación se lee como texto para no perder ceros ni formato
        dtype={"PRESTAMO": str, "NEGOCIACION": str},
    )
    df.columns = [str(c).strip().upper() for c in df.columns]   # limpia espacios en los encabezados
    faltan = requeridas - set(df.columns)                       # ¿falta alguna columna requerida?
    if faltan:
        raise ValueError(f"{ruta.name}: faltan columnas {sorted(faltan)}")
    # Deja las columnas en el orden definido y aplica los nombres finales
    return df[cfg["columnas"]].rename(columns=cfg["renombrar"])


def total_capital(serie: pd.Series) -> float:
    """Conteo de registros (valores no numéricos cuentan como 0).
    Se usa como control adicional además del conteo de registros."""
    return float(pd.to_numeric(serie, errors="coerce").fillna(0).sum())


# ------------------------------------------------------------------ PROCESO PRINCIPAL
def ejecutar(salida: Path) -> int:
    """Ejecuta todo el proceso. Devuelve 0 si terminó bien y 1 si hubo error."""
    log = logging.info      # atajo para mensajes informativos

    # ---- Fecha de referencia
    ref = (datetime.now() - timedelta(days=1)).date()       # datos al día anterior
    log("=== INICIO DEL PROCESO === Fecha de referencia: %s", ref)

    # ---- PASO 1: ubicar carpetas y archivos de origen
    log("[1/5] Buscando carpetas y archivos de origen...")
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
    log("    Carpeta de mes: %s", carpeta_mes)

    todos, descartados = localizar_archivos(carpeta_mes)    # busca en todas las carpetas de día
    for n in sorted(descartados):
        log("    Omitido por formato de nombre: %s", n)

    # Solo se consolida UN mes: se conservan los archivos del mismo año/mes de la fecha de referencia
    # y con fecha menor o igual a ella (nunca se mezclan meses)
    archivos = {k: v for k, v in todos.items()
                if (k[1].year, k[1].month) == (ref.year, ref.month) and k[1] <= ref}
    for tipo, fecha in sorted(set(todos) - set(archivos)):
        logging.warning("    Fuera del mes en proceso, se omite: %s %s", tipo, fecha)
    if not archivos:
        logging.error("No se encontraron archivos válidos en %s", carpeta_mes)
        return 1

    # ---- PASO 2: verificar que AMBOS archivos existan y estén a la misma fecha
    log("[2/5] Verificando que ambos archivos originales existan y estén a la misma fecha...")
    ultima = {}             # última fecha disponible de cada tipo
    for tipo, cfg in TIPOS.items():
        fechas_tipo = [f for (t, f) in archivos if t == tipo]
        if not fechas_tipo:                                 # no hay ningún archivo de este tipo
            logging.error("ERROR: no existe ningún archivo %s_ddmmyyyy en el mes %s. "
                          "Se aborta el proceso.", cfg["nombre"], MESES[ref.month - 1])
            return 1
        ultima[tipo] = max(fechas_tipo)
        log("    %s: última fecha disponible %s", cfg["nombre"], ultima[tipo].strftime("%d/%m/%Y"))

    if len(set(ultima.values())) != 1:                      # las fechas más recientes no coinciden
        detalle = " | ".join(f"{TIPOS[t]['nombre']} al {f.strftime('%d/%m/%Y')}" for t, f in ultima.items())
        logging.error("ERROR: los archivos originales NO están actualizados a la misma fecha (%s). "
                      "No se puede consolidar. Se aborta el proceso.", detalle)
        return 1

    fecha_datos = ultima["Saldos"]      # fecha hasta la que están actualizados los datos
    log("    OK: ambos archivos existen y están al %s", fecha_datos.strftime("%d/%m/%Y"))

    # Aviso si ambos archivos de esa fecha no están en la misma carpeta de día
    if archivos[("Saldos", fecha_datos)].parent != archivos[("Descuentos", fecha_datos)].parent:
        logging.warning("    Los archivos del %s están en carpetas distintas.", fecha_datos.strftime("%d/%m/%Y"))

    fechas = sorted({f for _, f in archivos})               # todas las fechas encontradas en el mes
    log("    Archivos válidos: %d | Rango: %s a %s",
        len(archivos), fechas[0].strftime("%d/%m/%Y"), fecha_datos.strftime("%d/%m/%Y"))

    # Informativo: días hábiles (lun-vie) sin archivo. Fines de semana no se reportan.
    d, faltantes = fechas[0].replace(day=1), []
    while d <= fecha_datos:
        if d.weekday() < 5 and d not in fechas:             # weekday < 5 = lunes a viernes
            faltantes.append(d.strftime("%d/%m"))
        d += timedelta(days=1)
    if faltantes:
        log("    Días hábiles sin archivo (feriados u omisiones): %s", ", ".join(faltantes))

    # Aviso si una fecha intermedia tiene un tipo de archivo pero le falta el otro
    for f in fechas:
        for tipo, cfg in TIPOS.items():
            if (tipo, f) not in archivos:
                logging.warning("    Falta %s para la fecha %s", cfg["nombre"], f.strftime("%d/%m/%Y"))

    # Base del nombre de salida: la fecha es la de los datos (no la de ejecución).
    # Cada parte se llamará <base_salida>_parteNN.csv
    base_salida = f"{PREFIJO_SALIDA}_{fecha_datos:%Y%m%d}"

    # ---- PASO 3: leer cada archivo
    log("[3/5] Leyendo archivos...")
    claves = sorted(archivos, key=lambda k: (k[1], k[0]))   # orden por fecha y luego por tipo
    partes, control = [], {}
    fechas_invalidas = set()

    for i, (tipo, fecha) in enumerate(claves, 1):
        ruta = archivos[(tipo, fecha)]
        if fecha in fechas_invalidas:
            continue
        log("    (%d/%d) Leyendo %s ...", i, len(claves), ruta.name)
        try:
            df = leer(ruta, tipo)
        except ValueError as e:
            logging.warning("%s", e)
            fechas_invalidas.add(fecha)
            partes = [
                p for p in partes
                if p["Fecha"].iloc[0].date() != fecha
            ]
            control = {
                k: v
                for k, v in control.items()
                if fecha.strftime("%d%m%Y") not in k
            }
            logging.warning(
                "Se omite completamente la fecha %s por columnas faltantes.",
                fecha.strftime("%d/%m/%Y")
            )
            continue

        control[ruta.name] = {
            "filas": len(df),
            "capital": total_capital(df["CAPITAL_TOTAL"])
        }
        df.insert(0, "Archivo", ruta.name)
        df.insert(0, "Fecha", pd.Timestamp(fecha))
        df.insert(0, "Documento", tipo)
        partes.append(df)
        log("        %d registros", len(df))

    if fechas_invalidas:
        logging.warning(
            "Fechas excluidas por estructura inválida: %s",
            ", ".join(f.strftime("%d/%m/%Y") for f in sorted(fechas_invalidas))
        )

    # ---- PASO 4: consolidar, eliminar los archivos anteriores (delete) y generar los CSV (insert)
    log("[4/5] Consolidando datos y generando CSV en partes (puede tardar)...")
    final = pd.concat(partes, ignore_index=True)            # apila todas las tablas en una sola
    del partes                                              # libera memoria
    final.insert(0, "No", range(1, len(final) + 1))         # correlativo de línea desde 1 (continúa entre partes)
    final = final[ORDEN_FINAL]                              # ordena las columnas

    # Delete: todos los archivos de la MISMA fecha (con o sin sufijo _parteNN, así también se
    # limpian partes sobrantes de una corrida anterior); los de otras fechas del mes, según configuración
    for p in sorted(salida.glob(f"{PREFIJO_SALIDA}_{fecha_datos:%Y%m}??*.csv")):
        misma_fecha = p.name.startswith(base_salida)
        if not misma_fecha and not ELIMINAR_ANTERIORES:
            continue
        try:
            p.unlink()                                      # elimina el archivo
            log("    Eliminado: %s", p.name)
        except PermissionError:                             # abierto en Excel u otro programa
            if misma_fecha:                                 # sin borrarlo no se puede recrear
                logging.error("No se pudo eliminar %s (¿está abierto?). Ciérrelo y reintente.", p.name)
                return 1
            logging.warning("    No se pudo eliminar %s (¿está abierto?). Se continúa.", p.name)

    # Insert: se escribe un CSV por cada bloque de FILAS_POR_PARTE registros.
    # utf-8-sig hace que Excel reconozca tildes y ñ; date_format escribe las fechas como dd/mm/yyyy
    total_generado = len(final)                             # filas que se escribieron
    n_partes = max(1, -(-total_generado // FILAS_POR_PARTE))   # división hacia arriba
    archivos_csv = []
    for n in range(n_partes):
        bloque = final.iloc[n * FILAS_POR_PARTE:(n + 1) * FILAS_POR_PARTE]
        ruta_parte = salida / f"{base_salida}_parte{n + 1:02d}.csv"
        bloque.to_csv(ruta_parte, index=False, encoding="utf-8-sig", date_format="%d/%m/%Y")
        archivos_csv.append(ruta_parte)
        log("    Parte %d/%d generada: %s (%s registros)",
            n + 1, n_partes, ruta_parte.name, f"{len(bloque):,}")
    del final

    # ---- PASO 5: validar contra los archivos ya escritos en disco
    log("[5/5] Validando registros entre archivos originales y CSV generados...")
    conteo, capital = {}, {}        # lo que realmente quedó en los CSV (sumando todas las partes)
    # Se relee cada CSV por bloques (chunks) para no cargar todo en memoria
    for ruta_parte in archivos_csv:
        for chunk in pd.read_csv(ruta_parte, usecols=["Archivo", "CAPITAL_TOTAL"],
                                 chunksize=500_000, encoding="utf-8-sig"):
            for nombre, g in chunk.groupby("Archivo"):
                conteo[nombre] = conteo.get(nombre, 0) + len(g)
                capital[nombre] = capital.get(nombre, 0.0) + total_capital(g["CAPITAL_TOTAL"])

    ok = True
    for nombre, c in control.items():                       # compara archivo por archivo
        if conteo.get(nombre, 0) != c["filas"] or abs(capital.get(nombre, 0.0) - c["capital"]) > 0.01:
            ok = False
            logging.error("DIFERENCIA en %s: origen=%d filas | CSV=%d filas",
                          nombre, c["filas"], conteo.get(nombre, 0))

    total_origen = sum(c["filas"] for c in control.values())    # total de registros en los originales
    total_csv = sum(conteo.values())                            # total de registros en los CSV
    if not (total_origen == total_generado == total_csv):       # los tres conteos deben coincidir
        ok = False

    log("Archivos procesados       : %d", len(control))
    log("Total registros originales: %s", f"{total_origen:,}")
    log("Total registros en CSV    : %s (en %d parte(s))", f"{total_csv:,}", len(archivos_csv))

    if not ok:
        logging.error("PROCESO TERMINADO CON ERRORES: la cantidad de registros NO coincide.")
        for ruta_parte in archivos_csv:                     # no se deja ninguna parte que no pasó la validación
            try:
                ruta_parte.unlink()
                logging.error("Se eliminó %s para que no sea utilizado.", ruta_parte.name)
            except OSError:
                logging.error("No se pudo eliminar %s; NO lo utilice.", ruta_parte.name)
        return 1

    log("Validación correcta: la cantidad de registros coincide.")
    log("Datos actualizados hasta el %s", fecha_datos.strftime("%d/%m/%Y"))
    log("=== PROCESO CONCLUIDO CON EXITO === %d archivo(s) en %s", len(archivos_csv), salida)
    return 0


def main() -> int:
    salida = Path(RUTA_SALIDA)
    salida.mkdir(parents=True, exist_ok=True)               # crea la carpeta de salida si no existe
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",     # cada línea lleva fecha/hora y nivel
        handlers=[logging.FileHandler(salida / NOMBRE_LOG, encoding="utf-8"),   # al archivo log
                  logging.StreamHandler(sys.stdout)],                           # y a la pantalla
    )
    try:
        return ejecutar(salida)
    except Exception:
        # Cualquier error inesperado (red caída, columna faltante, etc.) queda en pantalla y en el log
        logging.exception("PROCESO TERMINADO CON ERRORES (error no controlado)")
        return 1


# ------------------------------------------------------------------ PUNTO DE ENTRADA
if __name__ == "__main__":
    sys.exit(main())    # devuelve 0 (OK) o 1 (error)