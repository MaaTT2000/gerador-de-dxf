import os
import re
import math
import io
import zipfile
import json
from flask import Flask, render_template, request, send_file, jsonify
import ezdxf
import pandas as pd
from datetime import datetime

app = Flask(__name__)

#==============================================================================
# CONFIGURAÇÕES E CONSTANTES GLOBAIS
#==============================================================================
# Mapeamento de nomes de colunas amigáveis (planilha/formulário) para nomes internos
COLUMN_MAPPING = {
    # Chaves de identificação e forma
    'nome_arquivo': 'part_name',
    'custom_filename': 'part_name',
    'drawing_code': 'part_name',
    'forma': 'shape',
    # Chaves de dimensão
    'largura': 'width',
    'altura': 'height',
    'diametro': 'diameter',
    'base_(cateto_1)': 'rt_base',
    'altura_(cateto_2)': 'rt_height',
    'base': 'triangle_base',
    # Chaves de opções (Bloco de Texto e Cotas)
    'habilitar_bloco': 'include_text_info',
    'cotas': 'include_dims',  # <-- NOVA ADIÇÃO
    'qtd': 'part_quantity',
    'espessura': 'material_thickness',
}

#==============================================================================
# FUNÇÃO DE DESENHO (LÓGICA PURA)
#==============================================================================
def create_dxf_drawing(params: dict):
    """Gera um desenho DXF a partir de um dicionário de parâmetros já validado e preparado."""
    try:
        doc = ezdxf.new('R2000')
        msp = doc.modelspace()
        
        styles = params.get('styles', {})
        doc.layers.new('CONTORNO', dxfattribs={'color': styles.get('contour_color', 7)})
        doc.layers.new('FUROS', dxfattribs={'color': styles.get('holes_color', 1)})
        
        if params.get('text_lines'):
            doc.layers.new('TEXTO', dxfattribs={'color': styles.get('text_color', 2)})
            
        if styles.get('include_dims', False):
            doc.layers.new('COTAS', dxfattribs={'color': styles.get('text_color', 2)})
            doc.dimstyles.new('NOROACO_DIMSTYLE', dxfattribs={'dimtxt': styles.get('char_height', 5)})
            dim_attribs = {'layer': 'COTAS', 'dimstyle': 'NOROACO_DIMSTYLE'}

        shape_type = params.get('shape')
        shape_creators = {
            'rectangle': lambda m, p: m.add_lwpolyline([(0,0),(p['width'],0),(p['width'],p['height']),(0,p['height'])], close=True),
        }
        if shape_type in shape_creators:
            shape_creators[shape_type](msp, params).dxf.layer = 'CONTORNO'
        else:
            return None, f"Forma '{shape_type}' desconhecida."

        if params.get('text_lines'):
            start_point = styles.get('text_insert_point', (0, -20))
            char_height = styles.get('char_height', 5)
            line_spacing = char_height * 1.5
            for i, line in enumerate(params['text_lines']):
                y_pos = start_point[1] - (i * line_spacing)
                msp.add_text(
                    line,
                    dxfattribs={'layer': 'TEXTO', 'height': char_height, 'insert': (start_point[0], y_pos)}
                )

        if styles.get('include_dims'):
            dim_distance = styles.get('dim_distance', 20)
            dims_creators = {
                'rectangle': lambda m, p: (
                    m.add_aligned_dim(p1=(0, p['height']), p2=(p['width'], p['height']), distance=dim_distance, dxfattribs=dim_attribs).render(),
                    m.add_aligned_dim(p1=(0, 0), p2=(0, p['height']), distance=-dim_distance, dxfattribs=dim_attribs).render()
                ),
            }
            if shape_type in dims_creators:
                dims_creators[shape_type](msp, params)

        stream = io.StringIO()
        doc.write(stream)
        sanitized_filename = re.sub(r'[^\w.-]+', '_', str(params.get('part_name')))
        return stream.getvalue(), f"{sanitized_filename}.dxf"

    except KeyError as e:
        return None, f"Parâmetro obrigatório ausente para a forma '{shape_type}': {e}"
    except Exception as e:
        app.logger.error(f"Erro inesperado no desenho do DXF: {e}")
        return None, "Erro interno de desenho."

#==============================================================================
# FUNÇÃO CENTRAL DE PREPARAÇÃO E VALIDAÇÃO DE DADOS
#==============================================================================
def _prepare_data_for_dxf(raw_data: dict):
    """
    Recebe dados brutos (de formulário ou planilha), limpa, traduz, valida e prepara para desenho.
    """
    params = {}
    # 1. Traduz as chaves para o padrão interno
    for old_key, value in raw_data.items():
        clean_key = str(old_key).strip().lower().replace(' ', '_')
        internal_key = COLUMN_MAPPING.get(clean_key, clean_key)
        params[internal_key] = value

    # 2. Validação essencial
    if not params.get('part_name') or not params.get('shape'):
        return None, f"Dados insuficientes: 'part_name' ou 'shape' ausentes."

    # 3. Converte tipos de forma segura
    def to_float(value, default=0.0):
        if value is None: return default
        try:
            return float(str(value).replace(',', '.'))
        except (ValueError, TypeError):
            return default

    for key in ['width', 'height', 'diameter', 'material_thickness', 'material_density', 'part_quantity']:
        if key in params:
            params[key] = to_float(params[key], 1.0 if key == 'part_quantity' else 0.0)

    # 4. Cálculo de estilos dinâmicos
    max_dim = max(params.get('width', 0), params.get('height', 0), params.get('diameter', 0))
    max_dim = max_dim if max_dim > 0 else 200
    char_height = min(max(max_dim / 25, 5), 35)
    
    params['styles'] = {
        'contour_color': int(to_float(params.get('contour_color', 7))),
        'holes_color': int(to_float(params.get('holes_color', 1))),
        'text_color': int(to_float(params.get('text_color', 2))),
        'include_dims': str(params.get('include_dims', '')).lower() in ['true', 'on', '1', 'sim'],
        'char_height': char_height,
        'dim_distance': max(15, char_height * 3),
        'text_insert_point': (0, -char_height * 2),
    }
    
    # 5. Lógica do bloco de texto
    should_include_text = str(params.get('include_text_info', '')).lower() in ['true', 'on', '1', 'sim']
    if should_include_text:
        try:
            shape = params.get('shape')
            area_mm2 = 0
            if shape == 'rectangle': area_mm2 = params.get('width', 0) * params.get('height', 0)
            
            if area_mm2 <= 0: raise ValueError("Área da peça é zero.")

            volume_m3 = (area_mm2 / 1_000_000) * (params.get('material_thickness', 0) / 1_000)
            unit_weight_kg = volume_m3 * params.get('material_density', 7850)
            quantity = int(params.get('part_quantity', 1))
            total_weight_kg = unit_weight_kg * quantity

            params['text_lines'] = [
                f"{str(params.get('part_name', '')).upper()}",
                f"Espessura: {params.get('material_thickness', 0):.2f} mm  (Qtd: {quantity:02d}x)",
                f"Peso Unitario: {unit_weight_kg:.3f} Kg".replace('.', ','),
                f"Peso Total: {total_weight_kg:.3f} Kg".replace('.', ',')
            ]
        except Exception as e:
            app.logger.warning(f"Não foi possível calcular o peso para '{params.get('part_name')}': {e}")
            params['text_lines'] = [f"{str(params.get('part_name', '')).upper()}", "ERRO NO CALCULO DE PESO"]
            
    return params, None

#==============================================================================
# ROTAS FLASK (CONTROLADORES)
#==============================================================================
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate-batch', methods=['POST'])
def generate_dxf_batch():
    if 'spreadsheet_file' not in request.files: return "Nenhum arquivo de planilha foi enviado.", 400
    file = request.files['spreadsheet_file']
    if file.filename == '': return "Nenhum arquivo selecionado.", 400

    try:
        df = pd.read_excel(file, dtype=str).dropna(how='all')
    except Exception as e:
        app.logger.error(f"Erro ao ler a planilha: {e}")
        return "Erro ao processar o arquivo da planilha.", 500

    memory_file = io.BytesIO()
    with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
        for index, row in df.iterrows():
            raw_data = row.to_dict()
            raw_data.update(request.form)

            prepared_data, error = _prepare_data_for_dxf(raw_data)
            if error:
                app.logger.warning(f"Ignorando linha {index + 2}: {error}")
                continue

            dxf_content, filename = create_dxf_drawing(prepared_data)
            if error or not dxf_content:
                app.logger.warning(f"Falha ao desenhar a peça '{prepared_data.get('part_name')}': {filename or error}")
                continue
            
            zf.writestr(filename, dxf_content)

    memory_file.seek(0)
    zip_filename = f"LOTE_DXF_{datetime.now():%Ym%d_%H%M%S}.zip"
    return send_file(memory_file, mimetype='application/zip', as_attachment=True, download_name=zip_filename)

@app.route('/generate-dxf', methods=['POST'])
def generate_dxf_from_form():
    raw_data = request.form.to_dict(flat=True)
    prepared_data, error = _prepare_data_for_dxf(raw_data)
    if error:
        return f"Erro nos dados enviados: {error}", 400

    dxf_content, filename = create_dxf_drawing(prepared_data)
    if not dxf_content:
        return f"Erro ao gerar o desenho: {filename}", 500

    return send_file(io.BytesIO(dxf_content.encode()), as_attachment=True, download_name=filename, mimetype='application/dxf')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)