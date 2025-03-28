import os
import psycopg2
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Border, Alignment
from datetime import datetime
from dateutil.relativedelta import relativedelta
import sys

def main():
    # Leer las credenciales y la ruta del archivo desde variables de entorno
    db_name = os.environ.get('DB_NAME', 'semillas')
    db_user = os.environ.get('DB_USER', 'openerp')
    db_password = os.environ.get('DB_PASSWORD', '')
    db_host = os.environ.get('DB_HOST', '2.136.142.253')
    db_port = os.environ.get('DB_PORT', '5432')
    # Usar el archivo "Copìa_Portes.xlsx" (asegúrate de que esté en una ubicación persistente)
    file_path = os.environ.get('EXCEL_FILE_PATH', 'Copìa_Portes.xlsx')
    
    db_params = {
        'dbname': db_name,
        'user': db_user,
        'password': db_password,
        'host': db_host,
        'port': db_port
    }
    
    # Calcular fechas: desde el primer día de hace dos meses hasta el día actual.
    end_date = datetime.now()
    start_date = (end_date - relativedelta(months=2)).replace(day=1)
    end_date_str = end_date.strftime('%Y-%m-%d')
    start_date_str = start_date.strftime('%Y-%m-%d')
    
    # Consulta SQL (actualiza la consulta según tus necesidades)
    query = f"""
    SELECT DISTINCT ON (sm.id)
        sm.invoice_id as "ID FACTURA",
        rp.id as "ID CLIENTE",
        'S-' || rp.id AS "ID BBSeeds",
        sp.name as "DESCRIPCIÓN",
        sp.internal_number as "CÓDIGO FACTURA",
        sp.number as "NÚMERO DEL ASIENTO",
        to_char(sp.date_invoice, 'DD/MM/YYYY') as "FECHA FACTURA",
        sp.origin as "DOCUMENTO ORIGEN",
        pp.default_code as "REFERENCIA PRODUCTO", 
        pp.name as "NOMBRE", 
        COALESCE(pm.name,'') as "MARCA",
        s.name AS "SECCION", 
        f.name as "FAMILIA", 
        sf.name as "SUBFAMILIA",
        rc.name as "COMPAÑÍA",
        ssp.name as "SEDE",
        stp.date_preparado_app as "FECHA PREPARADO APP",
        (CASE WHEN stp.directo_cliente = true THEN 'Sí' ELSE 'No' END) AS "CAMIÓN DIRECTO",
        stp.number_of_packages as "NUMERO DE BULTOS",
        stp.num_pales as "NUMERO DE PALES",
        sp.portes as "PORTES",
        sp.portes_cubiertos as "PORTES CUBIERTOS",
        rp.nombre_comercial as "CLIENTE",
        rp.vat as "CIF CLIENTE",
        (CASE WHEN rpa.prov_id IS NOT NULL THEN (SELECT name FROM res_country_provincia WHERE id = rpa.prov_id) 
              ELSE rpa.state_id_2 END) as "PROVINCIA",
        rpa.city as "CIUDAD",
        (CASE WHEN rpa.cautonoma_id IS NOT NULL THEN (SELECT upper(name) FROM res_country_ca WHERE id = rpa.cautonoma_id) 
              ELSE '' END) as "COMUNIDAD",
        c.name as "PAÍS",
        to_char(sp.date_invoice, 'MM') as "MES",
        to_char(sp.date_invoice, 'DD') as "DÍA",
        spp.name as "PREPARADOR",
        sm.peso_arancel as "PESO",
        sum(CASE WHEN sp.type = 'out_invoice' THEN sm.cantidad_pedida
                 WHEN sp.type = 'out_refund' THEN -sm.cantidad_pedida END) as "UNIDADES VENTA",
        sum(CASE WHEN sp.type = 'out_invoice' THEN sm.price_subtotal
                 WHEN sp.type = 'out_refund' THEN -sm.price_subtotal END) as "BASE VENTA TOTAL",
        sum(CASE WHEN sp.type = 'out_invoice' THEN sm.margen
                 WHEN sp.type = 'out_refund' THEN -sm.margen END) as "MARGEN EUROS",
        sum(CASE WHEN sp.type = 'out_invoice' THEN sm.cantidad_pedida * sm.cost_price_real
                 WHEN sp.type = 'out_refund' THEN -sm.cantidad_pedida * sm.cost_price_real END) as "COSTE VENTA TOTAL"
    FROM account_invoice_line sm
    INNER JOIN account_invoice sp ON sp.id = sm.invoice_id
    INNER JOIN product_product pp ON sm.product_id = pp.id
    INNER JOIN res_partner_address rpa ON sp.address_invoice_id = rpa.id
    INNER JOIN res_country c ON c.id = rpa.pais_id
    LEFT OUTER JOIN stock_picking_invoice_rel spir ON spir.invoice_id = sp.id
    LEFT OUTER JOIN stock_picking stp ON stp.id = spir.pick_id
    LEFT OUTER JOIN stock_picking_preparador spp ON spp.id = stp.preparador
    LEFT OUTER JOIN res_company rc ON rc.id = sp.company_id
    LEFT OUTER JOIN res_partner rp ON rp.id = sp.partner_id
    LEFT OUTER JOIN stock_sede_ps ssp ON ssp.id = sp.sede_id
    LEFT OUTER JOIN product_category s ON s.id = pp.seccion
    LEFT OUTER JOIN product_category f ON f.id = pp.familia
    LEFT OUTER JOIN product_category sf ON sf.id = pp.subfamilia
    LEFT OUTER JOIN product_marca pm ON pm.id = pp.marca
    WHERE sp.type in ('out_invoice','out_refund')
      AND sp.state in ('open','paid')
      AND sp.journal_id != 11
      AND sp.anticipo = false
      AND pp.default_code NOT LIKE 'XXX%'
      AND sp.obsolescencia = false
      AND rp.nombre_comercial NOT LIKE '%PLANTASUR TRADING%'
      AND rp.nombre_comercial NOT LIKE '%PLANTADUCH%'
      AND sp.date_invoice >= '{start_date_str}'
      AND sp.date_invoice <= '{end_date_str}'
    GROUP BY 
        sm.id,
        rp.id,
        sm.company_id,
        sp.sede_id,
        sp.date_invoice,
        pp.seccion,
        pp.familia,
        pp.subfamilia,
        pp.default_code,
        pp.id,
        sm.cantidad_pedida,
        sp.partner_id,
        rpa.prov_id,
        rpa.state_id_2,
        rpa.cautonoma_id,
        c.name,
        sp.address_invoice_id,
        pp.proveedor_id1,
        sm.price_subtotal,
        sm.margen,
        pp.tarifa5,
        sp.directo_cliente,
        sp.obsolescencia,
        sp.anticipo,
        sp.name,
        s.name,
        f.name,
        sf.name,
        rc.name,
        ssp.name,
        stp.date_preparado_app,
        stp.directo_cliente,
        stp.number_of_packages,
        stp.num_pales,
        sp.portes,
        sp.portes_cubiertos,
        rp.nombre_comercial,
        spp.name,
        sp.internal_number,
        sp.origin,
        sp.number,
        rp.vat,
        pm.name,
        rpa.city
    ORDER BY 
        sm.id,
        stp.number_of_packages DESC;
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
    
    # Cargar el archivo Excel existente o crear uno nuevo si no existe
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
    
    # Evitar duplicados usando el ID (sm.id) como clave
    existing_ids = {row[0] for row in sheet.iter_rows(min_row=2, values_only=True)}
    for row in resultados:
        if row[0] not in existing_ids:
            # Añadir la fila nueva
            sheet.append(row)
            # Copiar el estilo de la última fila anterior (si existe) para mantener el formato
            new_row_index = sheet.max_row
            if new_row_index > 1:
                # Suponemos que la fila anterior tiene el formato deseado
                for col in range(1, sheet.max_column + 1):
                    source_cell = sheet.cell(row=new_row_index - 1, column=col)
                    target_cell = sheet.cell(row=new_row_index, column=col)
                    target_cell.font = source_cell.font
                    target_cell.fill = source_cell.fill
                    target_cell.border = source_cell.border
                    target_cell.alignment = source_cell.alignment
    
    book.save(file_path)
    print(f"Los datos se han guardado en el archivo {file_path}.")

if __name__ == '__main__':
    main()
