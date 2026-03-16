from flask import Flask, render_template_string, request, send_file
import pandas as pd
import json
import io

app = Flask(__name__)

HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Convertidor Total JSON</title>
    <style>
        body { font-family: sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; background: #f0f2f5; }
        .card { background: white; padding: 40px; border-radius: 10px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); text-align: center; }
        h2 { color: #1a73e8; }
        button { background: #1a73e8; color: white; border: none; padding: 12px 24px; cursor: pointer; border-radius: 4px; font-size: 16px; }
        button:hover { background: #1557b0; }
        input { margin-bottom: 20px; }
    </style>
</head>
<body>
    <div class="card">
        <h2>📂 Extractor Completo de JSON</h2>
        <p>Este motor extrae todos los niveles del archivo</p>
        <form method="post" enctype="multipart/form-data">
            <input type="file" name="file" accept=".json" required><br>
            <button type="submit">Convertir a Excel</button>
        </form>
    </div>
</body>
</html>
'''

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['file']
        if file:
            try:
                # 1. Leer el JSON
                data = json.load(file)
                
                # 2. Aplanado Total (Power Query Style)
                # sep='.' para que se parezca a lo que ves en Power Query
                df = pd.json_normalize(data, sep='.')

                # 3. TRUCO: Si el JSON es una sola factura, Excel la pone horizontal (larga).
                # Vamos a girarla automáticamente para que sea vertical y fácil de leer.
                if len(df) == 1:
                    df = df.T.reset_index()
                    df.columns = ['Campo/Ruta del JSON', 'Dato Extraído']

                # 4. Crear el archivo Excel sin florituras para que no dé error
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Datos')
                
                output.seek(0)
                return send_file(
                    output, 
                    download_name="extraccion_exitosa.xlsx", 
                    as_attachment=True,
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
                
            except Exception as e:
                # Si algo falla, te lo dirá en la pantalla
                return f"Hubo un detalle con este JSON específico: {str(e)}"
    
    return render_template_string(HTML_TEMPLATE)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)