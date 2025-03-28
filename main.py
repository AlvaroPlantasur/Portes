import os
import psycopg2
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from datetime import datetime
from dateutil.relativedelta import relativedelta
import sys
import copy

def main():
    # Obtener credenciales y ruta del archivo desde variables de entorno
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
    
    # Calcular rango de fechas: desde el primer día del mes de hace dos meses hasta el día actual.
    end_date = datetime.now()
    start_date = (end_date - relativedelta(months=2)).replace(day=1)
    end_date_str = end_date.strftime('%Y-%m-%d')
    start_date_str = start_date.strftime('%Y-%m-%d')
    
    # Consulta SQL (adaptada para usar el rango de fechas dinámico)
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
            WHEN rpa.prov_id IS NOT NULL THEN (SELECT UPPER(name) FROM res_country_provincia WHERE id = rpa.prov_id)
            ELSE UPPER(rpa.state_id_2)
        END AS "PROVINCIA",
        CASE 
            WHEN rpa.cautonoma_id IS NOT NULL THEN (SELECT UPPER(name) FROM res_country_ca WHERE id = rpa.cautonoma_id)
            ELSE ''
        END AS "COMUNIDAD",
        c.name AS "PAÍS",
        TO_CHAR(ai.date_invoice, 'MM') AS "MES",
        TO_CHAR(ai.date_invoice, 'DD') AS "DÍA",
        CASE 
            WHEN ai.type = 'out_invoice' THEN COALESCE(ai.portes, 0) + COALESCE(ai.portes_cubiertos, 0)
            ELSE -(COALESCE(ai.portes, 0) + COALESCE(ai.portes_cubiertos, 0))
        END AS "PORTES CARGADOS POR EL TRANSPORTISTA",
        CASE 
            WHEN ai.type = 'out_invoice' THEN COALESCE(ai.portes_cubiertos, 0)
            ELSE -(COALESCE(ai.portes_cubiertos, 0))
        END AS "PORTES CUBIERTOS",
        CASE 
            WHEN ai.type = 'out_invoice' THEN COALESCE(ai.portes, 0)
            ELSE -(COALESCE(ai.portes, 0))
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
    
    # Cargar el archivo Excel existente o crearlo si no existe
    try:
        book = load_workbook(file_path)
        sheet = book.active
    except FileNotFoundError:
        book = Workbook()
        sheet = book.active
        sheet.title = "Resultados"
        sheet.append(headers)
        for cell in sheet["1:1"]:
            cell.font = Font(bold=True)
    
    # Extraer los IDs ya existentes (se asume que la primera columna es "ID FACTURA")
    existing_ids = {row[0] for row in sheet.iter_rows(min_row=2, values_only=True)}
    
    # Añadir solo las filas nuevas para evitar duplicados
    for row in resultados:
        if row[0] not in existing_ids:
            sheet.append(row)
            # (Opcional) Puedes copiar estilos de la última fila si es necesario
            new_row_index = sheet.max_row
            if new_row_index > 1:
                for col in range(1, sheet.max_column + 1):
                    source_cell = sheet.cell(row=new_row_index - 1, column=col)
                    target_cell = sheet.cell(row=new_row_index, column=col)
                    target_cell.font = copy.copy(source_cell.font)
                    target_cell.fill = copy.copy(source_cell.fill)
                    target_cell.border = copy.copy(source_cell.border)
                    target_cell.alignment = copy.copy(source_cell.alignment)
    
    # Convertir el rango de datos en una tabla para Power BI
    max_row = sheet.max_row
    max_col = sheet.max_column
    last_col_letter = get_column_letter(max_col)
    table_ref = f"A1:{last_col_letter}{max_row}"
    
    if not sheet._tables:
        tab = Table(displayName="Table1", ref=table_ref)
        style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                               showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        tab.tableStyleInfo = style
        sheet.add_table(tab)
    else:
        for tab in sheet._tables:
            tab.ref = table_ref
    
    book.save(file_path)
    print(f"Los datos se han guardado en el archivo {file_path}.")

if __name__ == '__main__':
    main()
