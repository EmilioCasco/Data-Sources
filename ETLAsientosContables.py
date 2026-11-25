"""
=============================
CONSOLIDA ASIENTOS CONTABLES 
=============================
Fecha: 25/11/2026
Version: 1.0

Consolida los siguientes archivos:
Asientos Contables BDD
Asientos Contables BPR

Estos se encuentran en la ruta: //172.23.6.81/Listados2021 
"""

import logging                      # Para mostrar mensajes en pantalla y armar el log de la ejecución
import os                           # Para recorrer carpetas
import re                           # Expresiones regulares: validar el formato exacto del nombre
import sys                          # Para enviar la salida a pantalla y devolver el código de salida
import unicodedata                  # Para quitar tildes al comparar nombres de carpetas
from datetime import datetime, timedelta   # Manejo de fechas
from pathlib import Path            # Manejo cómodo de rutas de archivos

import pandas as pd                 # Lectura de Excel, consolidación y escritura de CSV

# ------------------------------------------------------------------ CONFIGURACIÓN
# Ruta de red donde están las carpetas por año (ruta UNC completa, con doble barra inicial)
RUTA_ORIGEN = r"C:\Users\HN39280\Desktop\Cartera\AsientosContables"
#RUTA_ORIGEN = r"\\172.23.6.81\Listados2021"


# Carpeta donde se guarda el CSV consolidado y el log
RUTA_SALIDA = r"C:\Users\HN39280\Desktop\Cartera\CP"

# Prefijo del archivo de salida: <PREFIJO_SALIDA>_yyyymmdd.csv
PREFIJO_SALIDA = "Consolidado_Asientos_Contables"
# Archivo de log (se guarda en la carpeta de salida)
NOMBRE_LOG = "AsientosContables_Log.log"

# True = al generar el nuevo CSV también se eliminan los consolidados de OTRAS fechas del
# mismo mes (así solo queda el más reciente). El de la misma fecha SIEMPRE se elimina y se
# vuelve a crear. Si algún archivo anterior está abierto en Excel, no se puede borrar: solo se avisa.
ELIMINAR_ANTERIORES = False

# Nombres de los meses tal como se llaman las carpetas (sin tildes y en minúscula,
# porque al comparar se normaliza el nombre real de la carpeta)
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

# Columnas que ambos archivos tienen en común (mismo nombre en los dos)
COLS_COMUNES = ["NUMERO_COMPROBANTE","NUMERO_TRANSACCION", "MONTO_COMPROBANTE","DESCRIPCION","MONTO_MONEDA_LOCAL",
               "CODIGO_EMPRESA","CODIGO_AGENCIA","CODIGO_SUB_APLICACION", "CODIGO_MONEDA","FECHA_MOVIMIENTO",
               "CODIGO_APLICACION","VALOR_CAMBIO","VALOR_CAMBIO_VAL","FECHA_VALIDA","CODIGO_CONCENTRADOR","TIPO_COMPROBANTE",
               "CODIGO_TIPO_TRANSACCION","CODIGO_TIPO_OPERACION","CODIGO_TIPO_SALDO","CODIGO_TIPO_SALDO_ANTERIOR",
               "NIVEL_1_CUENTA_CONT","NIVEL_2_CUENTA_CONT","NIVEL_3_CUENTA_CONT","NIVEL_4_CUENTA_CONT","NIVEL_5_CUENTA_CONT",
               "NIVEL_6_CUENTA_CONT","NIVEL_7_CUENTA_CONT","NIVEL_8_CUENTA_CONT"]

# Definición de cada tipo de archivo:
#   nombre    -> prefijo del archivo (solo se usa en los mensajes)
#   regex     -> formato EXACTO del nombre: prefijo + ddmmyyyy (8 dígitos) + extensión.
#   columnas  -> columnas que se leen del Excel
TIPOS = {
    "Asientos Contables BDD": {
        "nombre": "Asientos Contables BDD",
        "regex": re.compile(r"^Asientos Contables BDD(\d{8})\.(xlsx|xlsm|xls)$", re.I),
        "columnas": ["NUMERO_NEGOCIACION"] + COLS_COMUNES,
        # Se unifican los nombres para poder apilar ambos archivos en las mismas columnas
        "renombrar": {"NUMERO_NEGOCIACION": "NUMERO_PRESTAMO"},

    },
    "Asientos Contables BPR": {
        "nombre": "Asientos Contables BPR",
        "regex": re.compile(r"^Asientos Contables BPR(\d{8})\.(xlsx|xlsm|xls)$", re.I),
        "columnas": ["NUMERO_PRESTAMO"] + COLS_COMUNES,
        "renombrar": {"NUMERO_PRESTAMO": "NUMERO_PRESTAMO"},
    },
}


# Orden final de las columnas del CSV
ORDEN_FINAL = ["No","NUMERO_COMPROBANTE","NUMERO_TRANSACCION", "MONTO_COMPROBANTE","DESCRIPCION","MONTO_MONEDA_LOCAL",
               "CODIGO_EMPRESA","CODIGO_AGENCIA","CODIGO_SUB_APLICACION", "CODIGO_MONEDA","FECHA_MOVIMIENTO","CODIGO_APLICACION",
               "VALOR_CAMBIO","VALOR_CAMBIO_VAL","FECHA_VALIDA","CODIGO_CONCENTRADOR","TIPO_COMPROBANTE","NUMERO_PRESTAMO",
               "CODIGO_TIPO_TRANSACCION","CODIGO_TIPO_OPERACION","CODIGO_TIPO_SALDO","CODIGO_TIPO_SALDO_ANTERIOR",
               "NIVEL_1_CUENTA_CONT","NIVEL_2_CUENTA_CONT","NIVEL_3_CUENTA_CONT","NIVEL_4_CUENTA_CONT","NIVEL_5_CUENTA_CONT",
               "NIVEL_6_CUENTA_CONT","NIVEL_7_CUENTA_CONT","NIVEL_8_CUENTA_CONT","Documento", "Fecha", "Archivo"]

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
            if not coincide and nombre.lower().startswith(("asientos contables bdd", "asientos contables bpr")):
                descartados.append(nombre)
    return encontrados, descartados


def leer(ruta: Path, tipo: str) -> pd.DataFrame:
    """Lee un Excel y devuelve solo las columnas requeridas, ya con los nombres finales."""
    cfg = TIPOS[tipo]
    requeridas = {c.upper() for c in cfg["columnas"]}              # columnas esperadas, en mayúscula
    # Los .xls se leen con 'calamine' (xlrd está obsoleto y genera advertencias);
    # los .xlsx/.xlsm siguen con el motor por defecto (openpyxl)
    motor = "calamine" if ruta.suffix.lower() == ".xls" else None

    # BDD: solo primera hoja
    if tipo == "Asientos Contables BDD":
        df = pd.read_excel(
            ruta,
            engine=motor,
            usecols=lambda c: str(c).strip().upper() in requeridas,
            dtype=str
        )

        df.columns = [str(c).strip().upper() for c in df.columns]

        faltan = requeridas - set(df.columns)

        if faltan:
            raise ValueError(f"{ruta.name}: faltan columnas {sorted(faltan)}")

        return df[cfg["columnas"]].rename(columns=cfg["renombrar"])

# BPR: leer todas las hojas
    hojas = pd.read_excel(
        ruta,
        sheet_name=None,
        engine=motor,
        usecols=lambda c: str(c).strip().upper() in requeridas,
        dtype=str
    )

    dfs = []

    for nombre_hoja, df in hojas.items():
        if df.empty:
            continue

        df.columns = [str(c).strip().upper() for c in df.columns]
        faltan = requeridas - set(df.columns)

        if faltan:
            raise ValueError(
                f"{ruta.name} | Hoja '{nombre_hoja}': faltan columnas {sorted(faltan)}"
            )

        dfs.append(
            df[cfg["columnas"]].rename(columns=cfg["renombrar"])
        )

    if not dfs:
        raise ValueError(f"{ruta.name}: no contiene registros válidos")
    return pd.concat(dfs, ignore_index=True)

def total_registros(serie: pd.Series) -> float:
    """Conteo de NUMERO_COMPROBANTE (valores no numéricos cuentan como 0).
    Se usa como control adicional además del conteo de registros."""
    return float(pd.to_numeric(serie, errors="coerce").fillna(0).sum())

# ------------------------------------------------------------------ PROCESO PRINCIPAL
def ejecutar(salida: Path) -> int:
    """Ejecuta todo el proceso. Devuelve 0 si terminó bien y 1 si hubo error."""
    log = logging.info                                             # atajo para mensajes informativos

    # ---- Fecha de referencia
    ref = (datetime.now() - timedelta(days=1)).date()              # datos al día anterior
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

    # ---- PASO 2: verificar que AMBOS archivos existan y estén a la misma fecha
    log("[2/5] Verificando que ambos archivos originales existan y estén a la misma fecha...")
    ultima = {}                                                    # última fecha disponible de cada tipo
    for tipo, cfg in TIPOS.items():
        fechas_tipo = [f for (t, f) in archivos if t == tipo]
        if not fechas_tipo:                                        # no hay ningún archivo de este tipo
            logging.error("ERROR: no existe ningún archivo %s_ddmmyyyy en el mes %s. "
                          "Se aborta el proceso.", cfg["nombre"], MESES[ref.month - 1])
            return 1
        ultima[tipo] = max(fechas_tipo)
        log("      %s: última fecha disponible %s", cfg["nombre"], ultima[tipo].strftime("%d/%m/%Y"))
    if len(set(ultima.values())) != 1:                             # las fechas más recientes no coinciden
        detalle = " | ".join(f"{TIPOS[t]['nombre']} al {f.strftime('%d/%m/%Y')}" for t, f in ultima.items())
        logging.error("ERROR: los archivos originales NO están actualizados a la misma fecha (%s). "
                      "No se puede consolidar. Se aborta el proceso.", detalle)
        return 1
    fecha_datos = ultima["Asientos Contables BDD"]                                 # fecha hasta la que están actualizados los datos
    log("      OK: ambos archivos existen y están al %s", fecha_datos.strftime("%d/%m/%Y"))
    # Aviso si ambos archivos de esa fecha no están en la misma carpeta de día
    if archivos[("Asientos Contables BDD", fecha_datos)].parent != archivos[("Asientos Contables BPR", fecha_datos)].parent:
        logging.warning("      Los archivos del %s están en carpetas distintas.", fecha_datos.strftime("%d/%m/%Y"))

    fechas = sorted({f for _, f in archivos})                      # todas las fechas encontradas en el mes
    log("      Archivos válidos: %d | Rango: %s a %s",
        len(archivos), fechas[0].strftime("%d/%m/%Y"), fecha_datos.strftime("%d/%m/%Y"))
    # Informativo: días hábiles (lun-vie) sin archivo. Fines de semana no se reportan.
    d, faltantes = fechas[0].replace(day=1), []
    while d <= fecha_datos:
        if d.weekday() < 5 and d not in fechas:                    # weekday < 5 = lunes a viernes
            faltantes.append(d.strftime("%d/%m"))
        d += timedelta(days=1)
    if faltantes:
        log("      Días hábiles sin archivo (feriados u omisiones): %s", ", ".join(faltantes))
    # Aviso si una fecha intermedia tiene un tipo de archivo pero le falta el otro
    for f in fechas:
        for tipo, cfg in TIPOS.items():
            if (tipo, f) not in archivos:
                logging.warning("      Falta %s para la fecha %s", cfg["nombre"], f.strftime("%d/%m/%Y"))

    # Nombre del archivo de salida: la fecha es la de los datos (no la de ejecución)
    destino = salida / f"{PREFIJO_SALIDA}_{fecha_datos:%Y%m%d}.xlsx"

    # ---- PASO 3: leer cada archivo
    log("[3/5] Leyendo archivos...")
    claves = sorted(archivos, key=lambda k: (k[1], k[0]))          # orden por fecha y luego por tipo
    partes, control = [], {} 
    
    fechas_invalidas = set()

    for i, (tipo, fecha) in enumerate(claves, 1):
        ruta = archivos[(tipo, fecha)]
        
        if fecha in fechas_invalidas:
            continue

        log("      (%d/%d) Leyendo %s ...", i, len(claves), ruta.name)

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

        df.insert(0, "Archivo", ruta.name)
        df.insert(0, "Fecha", pd.Timestamp(fecha))
        df.insert(0, "Documento", tipo)

        partes.append(df)
        log("            %d registros", len(df))

    if fechas_invalidas:
        logging.warning(
            "Fechas excluidas por estructura inválida: %s",
            ", ".join(f.strftime("%d/%m/%Y") for f in sorted(fechas_invalidas))
    )

    # ---- PASO 4: consolidar, eliminar el archivo anterior (delete) y generar el CSV (insert)
    log("[4/5] Consolidando datos y generando XLS (puede tardar)...")
    final = pd.concat(partes, ignore_index=True)                   # apila todas las tablas en una sola
    # FILTRO NUEVO
    final = final[
        final["NIVEL_1_CUENTA_CONT"].astype(str).isin(["5011", "5012"])
        &
        (final["NIVEL_2_CUENTA_CONT"].astype(str) == "0301")
        &
        (final["NIVEL_3_CUENTA_CONT"].astype(str) == "0100")
    ].copy()

    control = {}

    for nombre, g in final.groupby("Archivo"):
        control[nombre] = {
            "filas": len(g),
            "capital": total_registros(g["NUMERO_COMPROBANTE"])
        }

    log("Registros después del filtro: %s", f"{len(final):,}")

    fechas = [
        "FECHA_MOVIMIENTO",
        "FECHA_VALIDA",
        "Fecha"
    ]

    for col in fechas:
        final[col] = (
            final[col]
            .astype(str)
            .str.replace(" 00:00:00", "", regex=False)
            .str.strip()
        )
        
        final[col] = pd.to_datetime(
            final[col],
            dayfirst=True,
            errors="coerce"
    )
    
    decimales = [
            "MONTO_COMPROBANTE",
            "MONTO_MONEDA_LOCAL",
            "VALOR_CAMBIO",
            "VALOR_CAMBIO_VAL"
            ]
    
    for col in decimales:
                final[col] = pd.to_numeric(final[col], errors="coerce")
    
    enteros = [
     
            "NUMERO_COMPROBANTE",
            "NUMERO_TRANSACCION",
            "CODIGO_EMPRESA",
            "CODIGO_AGENCIA",
            "CODIGO_SUB_APLICACION",
            "CODIGO_MONEDA",
            "CODIGO_APLICACION",
            "CODIGO_CONCENTRADOR",
            "NUMERO_PRESTAMO",
            "CODIGO_TIPO_TRANSACCION",
            "CODIGO_TIPO_OPERACION",
            "CODIGO_TIPO_SALDO",
            "CODIGO_TIPO_SALDO_ANTERIOR",
            "NIVEL_1_CUENTA_CONT",
            "NIVEL_2_CUENTA_CONT",
            "NIVEL_3_CUENTA_CONT",
            "NIVEL_4_CUENTA_CONT",
            "NIVEL_5_CUENTA_CONT",
            "NIVEL_6_CUENTA_CONT",
            "NIVEL_7_CUENTA_CONT",
            "NIVEL_8_CUENTA_CONT"
        ]
    
    for col in enteros:
            final[col] = pd.to_numeric(
                final[col],
                errors="coerce"
            ).astype("Int64")


    del partes                                                     # libera memoria
    final.insert(0, "No", range(1, len(final) + 1))                # correlativo de línea desde 1
    final = final[ORDEN_FINAL]                                     # ordena las columnas

    # Delete: el archivo con el MISMO nombre siempre se elimina; los de otras fechas del mes, según configuración
    for p in sorted(salida.glob(f"{PREFIJO_SALIDA}_{fecha_datos:%Y%m}??.xlsx")):
        if p != destino and not ELIMINAR_ANTERIORES:
            continue
        try:
            p.unlink()                                             # elimina el archivo
            log("      Eliminado: %s", p.name)
        except PermissionError:                                    # abierto en Excel u otro programa
            if p == destino:                                       # sin borrarlo no se puede recrear
                logging.error("No se pudo eliminar %s (¿está abierto?). Ciérrelo y reintente.", p.name)
                return 1
            logging.warning("      No se pudo eliminar %s (¿está abierto?). Se continúa.", p.name)

    # Insert: utf-8-sig hace que Excel reconozca tildes y ñ; date_format escribe las fechas como dd/mm/yyyy
    #final.to_csv(destino, index=False, encoding="utf-8-sig", date_format="%d/%m/%Y")

    with pd.ExcelWriter(destino, 
                        engine="openpyxl"
                        ) as writer:
        
        final.to_excel(
            writer,
            sheet_name="Comprobantes Contables",
            index=False
        )

        ws = writer.sheets["Comprobantes Contables"]

    # Decimales
        for columna in [
            "MONTO_COMPROBANTE",
            "MONTO_MONEDA_LOCAL",
            "VALOR_CAMBIO",
            "VALOR_CAMBIO_VAL"
        ]:
            idx = final.columns.get_loc(columna) + 1

            for col_cells in ws.iter_cols(
                min_col=idx,
                max_col=idx,
                min_row=2
            ):
            
                for cell in col_cells:
                    cell.number_format = "#,##0.00"
    # Fechas
        for col in ["FECHA_MOVIMIENTO", "FECHA_VALIDA", "Fecha"]:
            final[col] = (
                final[col]
                .astype(str)
                .str.replace(" 00:00:00", "", regex=False)
                .str.strip()
            )

            final[col] = pd.to_datetime(
                final[col],
                format="%m/%d/%Y",
                errors="coerce"
            )

            mascara = final[col].isna()

            final.loc[mascara, col] = pd.to_datetime(
                final.loc[mascara, col].astype(str),
                format="%d/%m/%Y",
                errors="coerce"
                )

        total_generado = len(final)
    log("      XLS generado: %s", destino) 

    # ---- PASO 5: validar contra el archivo ya escrito en disco
    log("[5/5] Validando registros entre archivos filtrados y XLS generado...")
    conteo, capital = {}, {}                                       # lo que realmente quedó en el CSV
    
    # Se relee el CSV por bloques (chunks) para no cargar todo en memoria
    validacion = pd.read_excel(
        destino,
        sheet_name="Comprobantes Contables",
        usecols=["Archivo", "NUMERO_COMPROBANTE"],
        dtype=str,
        engine="openpyxl"
    )
    
    for nombre, g in validacion.groupby("Archivo"):
        conteo[nombre] = len(g)
        capital[nombre] = total_registros(g["NUMERO_COMPROBANTE"])

    ok = True

    for nombre, c in control.items():                              # compara archivo por archivo
        if conteo.get(nombre, 0) != c["filas"] or abs(capital.get(nombre, 0.0) - c["capital"]) > 0.01:
            ok = False
            logging.error("DIFERENCIA en %s: origen=%d filas | XLS=%d filas",
                          nombre, c["filas"], conteo.get(nombre, 0))
    total_origen = sum(c["filas"] for c in control.values())       # total de registros en los originales
    total_xls = sum(conteo.values())                               # total de registros en el CSV
    if not (total_origen == total_generado == total_xls):          # los tres conteos deben coincidir
        ok = False

    log("Archivos procesados       : %d", len(archivos))
    log("Total registros originales filtrados: %s", f"{total_origen:,}")
    log("Total registros en .xlsx    : %s", f"{total_xls:,}")

    if not ok:
        logging.error("PROCESO TERMINADO CON ERRORES: la cantidad de registros NO coincide.")
        try:
            destino.unlink()                                       # no se deja un CSV que no pasó la validación
            logging.error("Se eliminó %s para que no sea utilizado.", destino.name)
        except OSError:
            logging.error("No se pudo eliminar %s; NO lo utilice.", destino.name)
        return 1

    log("Validación correcta: la cantidad de registros coincide.")
    log("Datos actualizados hasta el %s", fecha_datos.strftime("%d/%m/%Y"))
    log("=== PROCESO CONCLUIDO CON EXITO === %s", destino)
    return 0


def main() -> int:
    salida = Path(RUTA_SALIDA)
    salida.mkdir(parents=True, exist_ok=True)                      # crea la carpeta de salida si no existe
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",            # cada línea lleva fecha/hora y nivel
        handlers=[logging.FileHandler(salida / NOMBRE_LOG, encoding="utf-8"),  # al archivo log
                  logging.StreamHandler(sys.stdout)],              # y a la pantalla
    )
    try:
        return ejecutar(salida)
    except Exception:
        # Cualquier error inesperado (red caída, columna faltante, etc.) queda en pantalla y en el log
        logging.exception("PROCESO TERMINADO CON ERRORES (error no controlado)")
        return 1


# ------------------------------------------------------------------ PUNTO DE ENTRADA
if __name__ == "__main__":
    sys.exit(main())                                               # devuelve 0 (OK) o 1 (error)
