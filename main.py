import os
import psycopg2
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font
from datetime import datetime, timedelta

# Obtener datos sensibles desde variables de entorno
db_params = {
    'dbname': os.environ.get('DB_NAME', 'semillas'),
    'user': os.environ.get('DB_USER', 'openerp'),
    'password': os.environ.get('DB_PASSWORD', ''),  # Deja vacío o asigna un valor por defecto (no recomendado)
    'host': os.environ.get('DB_HOST', '2.136.142.253'),
    'port': os.environ.get('DB_PORT', '5432')
}

# La ruta donde se guardará el archivo Excel; se recomienda usar una ruta relativa
file_path = os.environ.get('EXCEL_FILE_PATH', 'hoy.xlsx')

# Calcular fechas (últimos 4 días)
end_date = datetime.now()
start_date = end_date - timedelta(days=4)

end_date_str = end_date.strftime('%Y-%m-%d')
start_date_str = start_date.strftime('%Y-%m-%d')

# Consulta SQL con fechas dinámicas
query = f"""
SELECT 
    ai.id AS "ID FACTURA",
    ai.date_invoice AS "FECHA FACTURA",
    ai.internal_number AS "CODIGO FACTURA",
    ai.name AS "DESCRIPCION",
    rc.name AS "COMPAÑÍA",
    ssp.name AS "SEDE",
    rp.nombre_comercial AS "CLIENTE",
    rpa.city AS "CIUDAD",
    (CASE WHEN rpa.prov_id IS NOT NULL THEN (SELECT UPPER(name) FROM res_country_provincia WHERE id = rpa.prov_id) 
        ELSE UPPER(rpa.state_id_2) 
    END) AS "PROVINCIA",
    (CASE WHEN rpa.cautonoma_id IS NOT NULL THEN (SELECT UPPER(name) FROM res_country_ca WHERE id = rpa.cautonoma_id) 
        ELSE '' 
    END) AS "COMUNIDAD",
    c.name AS "PAÍS",
    TO_CHAR(ai.date_invoice, 'MM') AS "MES",
    TO_CHAR(ai.date_invoice, 'DD') AS "DÍA",
    (CASE WHEN ai.type = 'out_invoice' THEN COALESCE(ai.portes,0) + COALESCE(ai.portes_cubiertos,0) 
    ELSE -(COALESCE(ai.portes,0) + COALESCE(ai.portes_cubiertos,0))
    END) AS "PORTES CARGADOS POR EL TRANSPORTISTA",
    (CASE WHEN ai.type = 'out_invoice' THEN COALESCE(ai.portes_cubiertos,0) 
    ELSE -(COALESCE(ai.portes_cubiertos,0))
    END) AS "PORTES CUBIERTOS",
    (CASE WHEN ai.type = 'out_invoice' THEN COALESCE(ai.portes,0) 
    ELSE -(COALESCE(ai.portes,0))
    END) AS "PORTES COBRADOS A CLIENTE"
FROM account_invoice ai
INNER JOIN res_partner_address rpa ON rpa.id = ai.address_shipping_id
INNER JOIN res_country c ON (c.id = rpa.pais_id)
LEFT JOIN stock_sede_ps ssp ON ssp.id = ai.sede_id
LEFT JOIN res_company rc ON rc.id = ai.company_id
LEFT JOIN res_partner rp ON rp.id = ai.partner_id
WHERE ai.state NOT IN ('draft','cancel') 
  AND ai.type IN ('out_invoice','out_refund') 
  AND ai.carrier_id IS NOT NULL 
  AND ai.date_invoice >= '{start_date_str}' 
  AND ai.date_invoice <= '{end_date_str}'
GROUP BY 
    ai.id,
    ai.company_id,
    ai.date_invoice,
    TO_CHAR(ai.date_invoice, 'YYYY'),
    TO_CHAR(ai.date_invoice, 'MM'),
    TO_CHAR(ai.date_invoice, 'YYYY-MM-DD'),
    ai.carrier_id,
    ai.partner_id,
    ai.name,
    ai.obsolescencia,
    ai.type,
    c.name,
    rpa.state_id_2,
    COALESCE(ai.portes,0) + COALESCE(ai.portes_cubiertos,0),
    COALESCE(ai.portes_cubiertos,0),
    COALESCE(ai.portes,0),
    rc.name,
    ssp.name,
    rpa.prov_id,
    rpa.cautonoma_id,
    rp.nombre_comercial,
    rpa.city
"""

# Conectar a la base de datos y ejecutar la consulta
try:
    with psycopg2.connect(**db_params) as conn:
        with conn.cursor() as cur:
            cur.execute(query)
            resultados = cur.fetchall()
            headers = [desc[0] for desc in cur.description]
except psycopg2.Error as e:
    print(f"Error al conectar o ejecutar la consulta: {e}")
    resultados = []

if not resultados:
    print("No se obtuvieron resultados de la consulta.")
else:
    print(f"Se obtuvieron {len(resultados)} filas de la consulta.")

# Intentar cargar el libro Excel existente para evitar duplicidades
try:
    book = load_workbook(file_path)
    sheet = book.active
except FileNotFoundError:
    # Crear un nuevo libro si no existe
    book = Workbook()
    sheet = book.active
    sheet.title = "Resultados"
    sheet.append(headers)
    for cell in sheet["1:1"]:
        cell.font = Font(bold=True)

# Crear un conjunto de IDs existentes para evitar duplicados
existing_ids = set()
for row in sheet.iter_rows(min_row=2, values_only=True):
    existing_ids.add(row[0])

# Añadir nuevas filas al archivo Excel si no están duplicadas
for row in resultados:
    if row[0] not in existing_ids:
        sheet.append(row)

# Guardar el archivo Excel
book.save(file_path)

print(f"Los datos de la consulta se han guardado en el archivo {file_path}.")
