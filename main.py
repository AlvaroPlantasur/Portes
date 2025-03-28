import os
import psycopg2
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from datetime import datetime
from dateutil.relativedelta import relativedelta
import sys
import copy

def main():
    # 1. Credenciales y archivo base
    db_name = os.environ.get('DB_NAME', 'semillas')
    db_user = os.environ.get('DB_USER', 'openerp')
    db_password = os.environ.get('DB_PASSWORD', '')
    db_host = os.environ.get('DB_HOST', '2.136.142.253')
    db_port = os.environ.get('DB_PORT', '5432')
    file_path = os.environ.get('EXCEL_FILE_PATH', 'Portes.xlsx')
    
    db_params = {
        'dbname': db_name,
        'user': db_user,
        'password': db_password,
        'host': db_host,
        'port': db_port
    }
    
    # 2. Calcular fechas
    end_date = datetime.now()
    start_date = (end_date - relativedelta(months=2)).replace(day=1)
    end_date_str = end_date.strftime('%Y-%m-%d')
    start_date_str = start_date.strftime('%Y-%m-%d')
    
    # 3. Consulta SQL
    query = f"""
    SELECT 
        ai.id AS "ID FACTURA",
        ai.date_invoice AS "FECHA FACTURA",
        ai.internal_number AS "CÓDIGO FACTURA",
        ai.name AS "DESCRIPCION",
        rc.name AS "COMPAÑÍA",
        ssp.name AS "SEDE",
        'S-' || rp.id AS "ID BBSeeds",
        rp.nombre_comercial AS "CLIENTE",
        rpa.city AS "CIUDAD",
        CASE 
            WHEN rpa.prov_id IS NOT NULL THEN 
                (SELECT UPPER(name) FROM res_country_provincia WHERE id = rpa.prov_id)
            ELSE 
                UPPER(rpa.state_id_2)
        END AS "PROVINCIA",
        CASE 
            WHEN rpa.cautonoma_id IS NOT NULL THEN 
                (SELECT UPPER(name) FROM res_country_ca WHERE id = rpa.cautonoma_id)
            ELSE 
                ''
        END AS "COMUNIDAD",
        c.name AS "PAÍS",
        TO_CHAR(ai.date_invoice, 'MM') AS "MES",
        TO_CHAR(ai.date_invoice, 'DD') AS "DÍA",
        CASE 
            WHEN ai.type = 'out_invoice' THEN 
                COALESCE(ai.portes, 0) + COALESCE(ai.portes_cubiertos, 0)
            ELSE 
                -(COALESCE(ai.portes, 0) + COALESCE(ai.portes_cubiertos, 0))
        END AS "PORTES CARGADOS POR EL TRANSPORTISTA",
        CASE 
            WHEN ai.type = 'out_invoice' THEN 
                COALESCE(ai.portes_cubiertos, 0)
            ELSE 
                -(COALESCE(ai.portes_cubiertos, 0))
        END AS "PORTES CUBIERTOS",
        CASE 
            WHEN ai.type = 'out_invoice' THEN 
                COALESCE(ai.portes, 0)
            ELSE 
                -(COALESCE(ai.portes, 0))
        END AS "PORTES COBRADOS A CLIENTE",
        TO_CHAR(ai.date_invoice, 'YYYY') AS "AÑO"
    FROM 
        account_invoice ai
    INNER JOIN 
        res_partner_address rpa ON rpa.id = ai.address_shipping_id
    INNER JOIN 
        res_country c ON c.id = rpa.pais_id
    LEFT JOIN 
        stock_sede_ps ssp ON ssp.id = ai.sede_id
    LEFT JOIN 
        res_company rc ON rc.id = ai.company_id
    LEFT JOIN 
        res_partner rp ON rp.id = ai.partner_id
    WHERE 
        ai.state NOT IN ('draft', 'cancel') 
        AND ai.type IN ('out_invoice', 'out_refund') 
        AND ai.carrier_id IS NOT NULL 
        AND ai.date_invoice BETWEEN '{start_date_str}' AND '{end_date_str}'
        AND ai.obsolescencia = FALSE
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
        COALESCE(ai.portes, 0) + COALESCE(ai.portes_cubiertos, 0),
        COALESCE(ai.portes_cubiertos, 0),
        COALESCE(ai.portes, 0),
        rc.name,
        ssp.name,
        rpa.prov_id,
        rpa.cautonoma_id,
        rp.nombre_comercial,
        rpa.city,
        rp.id
    ORDER BY 
        ai.id DESC;
    """
    
    # 4. Ejecutar consulta
    try:
        with psycopg2.connect(**db_params) as conn:
            with conn.cursor() as cur:
                cur.execute(query)
                resultados = cur.fetchall()
                headers = [desc[0] for desc in cur.description]
    except Exception as e:
        print(f"Error al conectar o ejecutar la consulta: {e}")
        sys.exit(1)
    
    if not resultados:
        print("No se obtuvieron resultados de la consulta.")
        return
    else:
        print(f"Se obtuvieron {len(resultados)} filas de la consulta.")
    
    # 5. Cargar el archivo base que ya tiene la tabla con el estilo
    try:
        book = load_workbook(file_path)
        sheet = book.active
    except FileNotFoundError:
        # Si no encuentra el archivo base con la tabla, no podrá reusar el estilo
        # Podrías crear uno nuevo, pero no tendrás el mismo estilo
        book = Workbook()
        sheet = book.active
        sheet.title = "Resultados"
        sheet.append(headers)
        for cell in sheet["1:1"]:
            cell.font = Font(bold=True)
    
    # 6. Evitar duplicados (asumiendo que la primera columna es "ID FACTURA")
    existing_ids = {row[0] for row in sheet.iter_rows(min_row=2, values_only=True)}
    for row in resultados:
        if row[0] not in existing_ids:
            sheet.append(row)
            # Si deseas copiar estilos de la última fila anterior, hazlo aquí
            # (similar a lo que tenías, con copy.copy de fonts, fill, etc.)
    
    # 7. Actualizar la referencia de la tabla existente
    # Asumiendo que la tabla en Excel se llama "MiTabla"
    # (Asegúrate de que en tu archivo base la tabla tenga ese nombre)
    if "MiTabla" in sheet.tables:
        tabla = sheet.tables["MiTabla"]
        max_row = sheet.max_row
        max_col = sheet.max_column
        last_col_letter = get_column_letter(max_col)
        new_ref = f"A1:{last_col_letter}{max_row}"
        tabla.ref = new_ref
    else:
        print("No se encontró la tabla 'MiTabla'. Se conservará el formato pero no se actualizará la referencia.")
    
    book.save(file_path)
    print(f"Archivo guardado con la estructura de tabla original en {file_path}.")

if __name__ == '__main__':
    main()
