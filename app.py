from flask import Flask, render_template, request, redirect, url_for, jsonify, send_file
import sqlite3
from datetime import datetime
import os
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, numbers

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'dinero-app-2024')
DB = os.environ.get('DB_PATH', 'dinero.db')


def get_db():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with get_db() as conn:
        conn.executescript('''
            CREATE TABLE IF NOT EXISTS rondas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL,
                total_inicial REAL NOT NULL DEFAULT 0,
                total_ganado REAL NOT NULL DEFAULT 0,
                fecha TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS participantes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ronda_id INTEGER NOT NULL,
                nombre TEXT NOT NULL,
                porcentaje REAL NOT NULL,
                FOREIGN KEY (ronda_id) REFERENCES rondas(id)
            );
        ''')


init_db()


def calcular_tabla(ronda, participantes):
    tabla = []
    for p in participantes:
        capital = ronda['total_inicial'] * p['porcentaje'] / 100
        ganado = ronda['total_ganado'] * p['porcentaje'] / 100
        total_queda = capital + ganado
        tabla.append({
            'nombre': p['nombre'],
            'porcentaje': p['porcentaje'],
            'capital_dado': capital,
            'dinero_ganado': ganado,
            'total_queda': total_queda,
            'ganancia': ganado,
        })
    return tabla


@app.route('/')
def index():
    conn = get_db()
    rondas = conn.execute('SELECT * FROM rondas ORDER BY fecha DESC').fetchall()
    conn.close()
    return render_template('index.html', rondas=rondas)


@app.route('/nueva', methods=['GET', 'POST'])
def nueva_ronda():
    if request.method == 'POST':
        nombre = request.form.get('nombre', '').strip()
        total_inicial = float(request.form.get('total_inicial', 0))
        total_ganado = float(request.form.get('total_ganado', 0))
        nombres_p = request.form.getlist('nombre_p')
        porcentajes = request.form.getlist('porcentaje')
        fecha = datetime.now().strftime('%Y-%m-%d %H:%M')

        if not nombre or not nombres_p:
            return render_template('nueva.html', error='Completa todos los campos.')

        conn = get_db()
        cur = conn.execute(
            'INSERT INTO rondas (nombre, total_inicial, total_ganado, fecha) VALUES (?, ?, ?, ?)',
            (nombre, total_inicial, total_ganado, fecha)
        )
        ronda_id = cur.lastrowid
        for n, p in zip(nombres_p, porcentajes):
            if n.strip() and p:
                conn.execute(
                    'INSERT INTO participantes (ronda_id, nombre, porcentaje) VALUES (?, ?, ?)',
                    (ronda_id, n.strip(), float(p))
                )
        conn.commit()
        conn.close()
        return redirect(url_for('ver_ronda', id=ronda_id))

    return render_template('nueva.html', error=None)


@app.route('/ronda/<int:id>')
def ver_ronda(id):
    conn = get_db()
    ronda = conn.execute('SELECT * FROM rondas WHERE id = ?', (id,)).fetchone()
    if not ronda:
        conn.close()
        return redirect(url_for('index'))
    participantes = conn.execute(
        'SELECT * FROM participantes WHERE ronda_id = ? ORDER BY porcentaje DESC', (id,)
    ).fetchall()
    conn.close()
    tabla = calcular_tabla(ronda, participantes)
    return render_template('ronda.html', ronda=ronda, tabla=tabla)


@app.route('/ronda/<int:id>/eliminar', methods=['POST'])
def eliminar_ronda(id):
    conn = get_db()
    conn.execute('DELETE FROM participantes WHERE ronda_id = ?', (id,))
    conn.execute('DELETE FROM rondas WHERE id = ?', (id,))
    conn.commit()
    conn.close()
    return redirect(url_for('index'))


@app.route('/resumen')
def resumen():
    conn = get_db()
    datos = conn.execute('''
        SELECT
            p.nombre,
            COUNT(DISTINCT p.ronda_id) AS rondas,
            SUM(r.total_inicial * p.porcentaje / 100) AS total_capital,
            SUM(r.total_ganado  * p.porcentaje / 100) AS total_ganado,
            SUM((r.total_inicial + r.total_ganado) * p.porcentaje / 100) AS total_recibido
        FROM participantes p
        JOIN rondas r ON p.ronda_id = r.id
        GROUP BY LOWER(TRIM(p.nombre))
        ORDER BY total_ganado DESC
    ''').fetchall()
    conn.close()
    return render_template('resumen.html', datos=datos)


@app.route('/exportar')
def exportar_excel():
    conn = get_db()
    rondas = conn.execute('SELECT * FROM rondas ORDER BY fecha DESC').fetchall()
    participantes = conn.execute('SELECT * FROM participantes').fetchall()
    conn.close()

    # Mapear participantes por ronda
    part_por_ronda = {}
    for p in participantes:
        part_por_ronda.setdefault(p['ronda_id'], []).append(p)

    wb = openpyxl.Workbook()

    # ── Hoja 1: Rondas ──────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = 'Rondas'
    header_fill = PatternFill('solid', fgColor='212529')
    header_font = Font(bold=True, color='FFFFFF')

    headers1 = ['Ronda', 'Fecha', 'Total Inicial ($)', 'Total Ganado ($)', 'Ganancia ($)']
    for col, h in enumerate(headers1, 1):
        cell = ws1.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')

    for r in rondas:
        ganancia = r['total_ganado'] - r['total_inicial']
        ws1.append([r['nombre'], r['fecha'], r['total_inicial'], r['total_ganado'], ganancia])

    for col in ws1.iter_cols(min_row=2, min_col=3, max_col=5):
        for cell in col:
            cell.number_format = '"$"#,##0'

    ws1.column_dimensions['A'].width = 25
    ws1.column_dimensions['B'].width = 18
    for col_letter in ['C', 'D', 'E']:
        ws1.column_dimensions[col_letter].width = 18

    # ── Hoja 2: Detalle por ronda ────────────────────────────────────
    ws2 = wb.create_sheet('Detalle')
    headers2 = ['Ronda', 'Fecha', 'Persona', 'Porcentaje (%)', 'Capital dado ($)', 'Ganado ($)', 'Total ($)']
    for col, h in enumerate(headers2, 1):
        cell = ws2.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')

    for r in rondas:
        for p in part_por_ronda.get(r['id'], []):
            capital = r['total_inicial'] * p['porcentaje'] / 100
            ganado  = r['total_ganado']  * p['porcentaje'] / 100
            total   = capital + ganado
            ws2.append([r['nombre'], r['fecha'], p['nombre'], p['porcentaje'], capital, ganado, total])

    for col in ws2.iter_cols(min_row=2, min_col=5, max_col=7):
        for cell in col:
            cell.number_format = '"$"#,##0'

    ws2.column_dimensions['A'].width = 25
    ws2.column_dimensions['B'].width = 18
    ws2.column_dimensions['C'].width = 20
    ws2.column_dimensions['D'].width = 16
    for col_letter in ['E', 'F', 'G']:
        ws2.column_dimensions[col_letter].width = 18

    # ── Hoja 3: Resumen por persona ──────────────────────────────────
    ws3 = wb.create_sheet('Resumen por persona')
    headers3 = ['Persona', 'Rondas', 'Capital total ($)', 'Total ganado ($)', 'Total recibido ($)', 'Ganancia neta ($)', 'ROI (%)']
    for col, h in enumerate(headers3, 1):
        cell = ws3.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')

    conn2 = get_db()
    resumen = conn2.execute('''
        SELECT
            p.nombre,
            COUNT(DISTINCT p.ronda_id) AS rondas,
            SUM(r.total_inicial * p.porcentaje / 100) AS total_capital,
            SUM(r.total_ganado  * p.porcentaje / 100) AS total_ganado,
            SUM((r.total_inicial + r.total_ganado) * p.porcentaje / 100) AS total_recibido
        FROM participantes p
        JOIN rondas r ON p.ronda_id = r.id
        GROUP BY LOWER(TRIM(p.nombre))
        ORDER BY total_ganado DESC
    ''').fetchall()
    conn2.close()

    for d in resumen:
        ganancia_neta = d['total_recibido'] - d['total_capital']
        roi = (ganancia_neta / d['total_capital'] * 100) if d['total_capital'] else 0
        ws3.append([d['nombre'], d['rondas'], d['total_capital'], d['total_ganado'], d['total_recibido'], ganancia_neta, roi])

    for col in ws3.iter_cols(min_row=2, min_col=3, max_col=6):
        for cell in col:
            cell.number_format = '"$"#,##0'
    for cell in ws3['G'][1:]:
        cell.number_format = '0.0"%"'

    ws3.column_dimensions['A'].width = 20
    ws3.column_dimensions['B'].width = 10
    for col_letter in ['C', 'D', 'E', 'F', 'G']:
        ws3.column_dimensions[col_letter].width = 20

    # Enviar archivo
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    filename = f'kuazimides_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx'
    return send_file(buf, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


if __name__ == '__main__':
    app.run(debug=True)
