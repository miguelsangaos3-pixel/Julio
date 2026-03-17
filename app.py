from flask import Flask, render_template_string, request, send_file
import pandas as pd
import json
import io

app = Flask(__name__)

# Campos específicos de los productos
CAMPOS_PRODUCTO = ['numItem', 'cantidad', 'descripcion', 'precioUni', 'montoDescu', 'ventaGravada']

HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Extractor Contable</title>
    <style>
        body { font-family: 'Segoe UI', sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; background: #eef2f3; }
        .card { background: white; padding: 40px; border-radius: 15px; box-shadow: 0 10px 25px rgba(0,0,0,0.1); text-align: center; border-top: 10px solid #facc15; }
        h2 { color: #854d0e; }
        button { background: #facc15; color: #854d0e; border: none; padding: 12px 25px; cursor: pointer; border-radius: 8px; font-weight: bold; }
        button:hover { background: #fde047; }
    </style>
</head>
<body>
    <div class="card">
        <h2>📑 Extractor de Json </h2>
        <p>Estraer Json</p>
        <form method="post" enctype="multipart/form-data">
            <input type="file" name="file" accept=".json" multiple required><br>
            <button type="submit">Generar Reporte</button>
        </form>
    </div>
</body>
</html>
'''

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        files = request.files.getlist('file')
        if files:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for idx, file in enumerate(files):
                    try:
                        data = json.load(file)
                        
                        # 1. Título Dinámico
                        tipo_dte = data.get('identificacion', {}).get('tipoDte', '00')
                        nombre_doc = "COMPROBANTE DE CRÉDITO FISCAL" if tipo_dte == '03' else "FACTURA CONSUMIDOR FINAL"
                        
                        # 2. SECCIÓN: DATOS GENERALES (Identificación, Emisor, Receptor)
                        # Combinamos las ramas principales para no perder nada de lo amarillo
                        gen_data = {
                            'identificacion': data.get('identificacion', {}),
                            'emisor': data.get('emisor', {}),
                            'receptor': data.get('receptor', {}),
                            'respuestaMH': data.get('responseMH', {})
                        }
                        df_general = pd.json_normalize(gen_data, sep='.').T.reset_index()
                        df_general.columns = ['Campo', 'Valor']

                        # 3. SECCIÓN: PRODUCTOS (Cuerpo del Documento)
                        cuerpo = data.get('cuerpoDocumento', [])
                        df_productos = pd.DataFrame(cuerpo)
                        if not df_productos.empty:
                            columnas_reales = [c for c in CAMPOS_PRODUCTO if c in df_productos.columns]
                            df_productos = df_productos[columnas_reales]

                        # 4. SECCIÓN: TOTALES (Resumen)
                        df_resumen = pd.json_normalize(data.get('resumen', {}), sep='.').T.reset_index()
                        df_resumen.columns = ['Campo', 'Valor']

                        # --- CONSTRUCCIÓN DE LA HOJA ---
                        sheet_name = f"Doc_{idx+1}"
                        
                        # Encabezado Principal
                        encabezado = pd.DataFrame([['TIPO DOCUMENTO', nombre_doc], ['ORIGEN', file.filename], ['='*30, '='*30]], columns=['Campo', 'Valor'])
                        encabezado.to_excel(writer, index=False, sheet_name=sheet_name, startrow=0)
                        
                        # Escribir Datos Generales (Emisor, Receptor, MH)
                        punto_productos = 5 + len(df_general)
                        df_general.to_excel(writer, index=False, sheet_name=sheet_name, startrow=4)
                        
                        # Escribir Productos
                        pd.DataFrame([['--- DETALLE DE PRODUCTOS ---', '']]).to_excel(writer, index=False, header=False, sheet_name=sheet_name, startrow=punto_productos + 1)
                        df_productos.to_excel(writer, index=False, sheet_name=sheet_name, startrow=punto_productos + 2)
                        
                        # Escribir Totales
                        punto_resumen = punto_productos + len(df_productos) + 4
                        pd.DataFrame([['--- RESUMEN Y TOTALES ---', '']]).to_excel(writer, index=False, header=False, sheet_name=sheet_name, startrow=punto_resumen)
                        df_resumen.to_excel(writer, index=False, header=False, sheet_name=sheet_name, startrow=punto_resumen + 1)

                    except Exception as e:
                        continue

            output.seek(0)
            return send_file(output, download_name="reporte_contable_completo.xlsx", as_attachment=True)
            
    return render_template_string(HTML_TEMPLATE)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)