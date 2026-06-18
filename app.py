import os
import re
import tempfile
from collections import defaultdict
from flask import Flask, render_template, request, jsonify, send_file
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024


def parse_num(s):
    if s is None:
        return None
    s = str(s).strip().replace('\xa0', '').replace(' ', '')
    try:
        return int(s)
    except (ValueError, TypeError):
        return None


def words_to_rows(words, y_tolerance=3):
    """Group words by their vertical position into rows."""
    rows = defaultdict(list)
    for w in words:
        # Round y to group nearby words
        y = round(w['top'] / y_tolerance) * y_tolerance
        rows[y].append(w)
    # Sort rows by y, words within each row by x
    result = []
    for y in sorted(rows.keys()):
        row_words = sorted(rows[y], key=lambda w: w['x0'])
        result.append((y, row_words))
    return result


def words_in_band(row_words, x_min, x_max):
    """Get concatenated text of words within x range."""
    tokens = [w['text'] for w in row_words if x_min <= w['x0'] < x_max]
    return ''.join(tokens)  # join without space for numbers (handles "1 000" → "1000")


def detect_document_type(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text() or ''
    if 'SOLDES AUX COMPTES' in text.upper():
        return 'soldes'
    elif 'VALORISATION PORTEFEUILLE' in text.upper():
        return 'valorisation'
    return 'unknown'


# ── Column boundaries (calibrated from word-position analysis) ─────────
# Soldes aux Comptes (x positions in PDF points):
SAC_TITRE_MAX = 200
SAC_ISIN_MIN = 200
SAC_ISIN_MAX = 330
SAC_DISPO_MIN = 330
SAC_DISPO_MAX = 410   # Vente starts at ~417
SAC_VENTE_MIN = 410
SAC_VENTE_MAX = 450
SAC_GELE_MIN = 450
SAC_GELE_MAX = 530
SAC_SOLDE_IND_MIN = 530

# Valorisation Portefeuille (x positions in PDF points):
VP_COMPTE_MAX = 210
VP_NCOMPTE_MIN = 210
VP_NCOMPTE_MAX = 260
VP_TITRE_MIN = 260
VP_TITRE_MAX = 365   # Cours can start at ~371
VP_COURS_MIN = 365
VP_COURS_MAX = 440
VP_BALANCE_MIN = 440
VP_BALANCE_MAX = 490  # Valo can start at ~498
VP_VALO_MIN = 490

ISIN_RE = re.compile(r'^[A-Z]{2}[0-9]{10}$')


def parse_soldes_aux_comptes(pdf_path):
    accounts = {}
    current_adherent = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ''

            # Detect adherent header from full text
            for line in text.split('\n'):
                m = re.match(r'Adh[eé]rent\s*-\s*(\d+)\s+(.+?)\s+\d{2}/\d{2}/\d{4}', line)
                if m:
                    adh_id = m.group(1).strip()
                    adh_name = m.group(2).strip()
                    current_adherent = adh_id
                    if adh_id not in accounts:
                        accounts[adh_id] = {
                            'adherent': adh_id,
                            'name': adh_name,
                            'lines': []
                        }

            if not current_adherent:
                continue

            words = page.extract_words()
            rows = words_to_rows(words, y_tolerance=4)

            for y, row_words in rows:
                # Check if this row has an ISIN
                isin_words = [w for w in row_words
                              if SAC_ISIN_MIN <= w['x0'] < SAC_ISIN_MAX
                              and ISIN_RE.match(w['text'])]
                if not isin_words:
                    continue

                isin = isin_words[0]['text']

                # Titre: all words left of ISIN column
                titre_words = [w['text'] for w in row_words if w['x0'] < SAC_TITRE_MAX]
                titre = ' '.join(titre_words).strip()

                # Numeric columns: join digit tokens within each band
                dispo = parse_num(words_in_band(row_words, SAC_DISPO_MIN, SAC_DISPO_MAX))
                vente = parse_num(words_in_band(row_words, SAC_VENTE_MIN, SAC_VENTE_MAX))
                gele = parse_num(words_in_band(row_words, SAC_GELE_MIN, SAC_GELE_MAX))
                solde_ind = parse_num(words_in_band(row_words, SAC_SOLDE_IND_MIN, 9999))

                accounts[current_adherent]['lines'].append({
                    'titre': titre,
                    'isin': isin,
                    'solde_disponible': dispo or 0,
                    'vente_en_attente': vente or 0,
                    'gele': gele or 0,
                    'nanti': 0,
                    'solde_indicatif': solde_ind or dispo or 0,
                })

    return list(accounts.values())


def parse_valorisation_portefeuille(pdf_path):
    accounts = {}

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ''
            words = page.extract_words()
            rows = words_to_rows(words, y_tolerance=4)

            # Get total from text
            m_total = re.search(r'Ligne de portefeuille\s*:\s*\d+\s+([\d\s]+)', text)
            page_total = None
            if m_total:
                page_total = parse_num(m_total.group(1).replace(' ', ''))

            last_account = None
            for y, row_words in rows:
                # Detect N° Compte: 6-digit number anywhere in the compte/ncompte area
                # Also handle concatenated case like "HALAL112464"
                n_compte = None
                compte_name_words = []
                for w in row_words:
                    m_nc = re.search(r'(\d{6})', w['text'])
                    if m_nc and w['x0'] < VP_TITRE_MIN:
                        n_compte = m_nc.group(1)
                        # Everything before this word (and before the number in this word)
                        for pw in row_words:
                            if pw is w:
                                break
                            if pw['x0'] < VP_COMPTE_MAX:
                                compte_name_words.append(pw['text'])
                        break

                if not n_compte:
                    continue

                compte_name = ' '.join(compte_name_words).strip()
                if not compte_name:
                    # Try text before the 6-digit number in the same word
                    for w in row_words:
                        m_nc = re.search(r'(\d{6})', w['text'])
                        if m_nc:
                            prefix = w['text'][:m_nc.start()].strip()
                            if prefix:
                                compte_name = prefix
                            break

                # Titre: words in titre band
                titre_words = [w['text'] for w in row_words
                               if VP_TITRE_MIN <= w['x0'] < VP_TITRE_MAX]
                titre = ' '.join(titre_words).strip()

                if not titre:
                    continue

                # Numeric columns
                cours = parse_num(words_in_band(row_words, VP_COURS_MIN, VP_COURS_MAX))
                balance = parse_num(words_in_band(row_words, VP_BALANCE_MIN, VP_BALANCE_MAX))
                valo = parse_num(words_in_band(row_words, VP_VALO_MIN, 9999))

                if n_compte not in accounts:
                    accounts[n_compte] = {
                        'adherent': n_compte,
                        'name': compte_name,
                        'total_valorisation': 0,
                        'lines': []
                    }
                accounts[n_compte]['lines'].append({
                    'titre': titre,
                    'cours_reference': cours,
                    'balance': balance or 0,
                    'valorisation': valo or 0,
                })
                last_account = n_compte

            if page_total and last_account:
                accounts[last_account]['total_valorisation'] = page_total

    # Recompute missing totals
    for key in accounts:
        if not accounts[key]['total_valorisation']:
            accounts[key]['total_valorisation'] = sum(
                l['valorisation'] for l in accounts[key]['lines']
            )

    return list(accounts.values())


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/parse', methods=['POST'])
def parse_pdf():
    if 'files' not in request.files:
        return jsonify({'error': 'No files provided'}), 400

    files = request.files.getlist('files')
    results = {'soldes': [], 'valorisation': [], 'errors': []}

    for f in files:
        if not f.filename.lower().endswith('.pdf'):
            results['errors'].append(f'{f.filename}: not a PDF')
            continue

        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
            f.save(tmp.name)
            tmp_path = tmp.name

        try:
            doc_type = detect_document_type(tmp_path)
            if doc_type == 'soldes':
                accounts = parse_soldes_aux_comptes(tmp_path)
                results['soldes'].extend(accounts)
            elif doc_type == 'valorisation':
                accounts = parse_valorisation_portefeuille(tmp_path)
                results['valorisation'].extend(accounts)
            else:
                results['errors'].append(f'{f.filename}: type de document non reconnu')
        except Exception as e:
            import traceback
            results['errors'].append(f'{f.filename}: {str(e)}\n{traceback.format_exc()}')
        finally:
            os.unlink(tmp_path)

    return jsonify(results)


@app.route('/api/export', methods=['POST'])
def export_excel():
    data = request.get_json()
    if not data:
        return jsonify({'error': 'No data'}), 400

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    def hfill(color):
        return PatternFill(start_color=color, end_color=color, fill_type='solid')

    header_fill = hfill('003366')
    alt_fill = hfill('E8F0FE')
    total_fill = hfill('FFF2CC')
    ok_fill = hfill('D4EDDA')
    err_fill = hfill('F8D7DA')

    thin = Side(style='thin', color='CCCCCC')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    num_fmt = '#,##0'

    def set_header(cell):
        cell.fill = header_fill
        cell.font = Font(color='FFFFFF', bold=True, size=10)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border

    def set_cell(cell, row_i, is_num=False):
        cell.fill = alt_fill if row_i % 2 == 0 else PatternFill()
        cell.border = border
        cell.alignment = Alignment(vertical='center',
                                   horizontal='right' if is_num else 'left')
        if is_num:
            cell.number_format = num_fmt

    # ── Soldes aux Comptes ─────────────────────────────────────────────
    soldes = data.get('soldes', [])
    if soldes:
        ws = wb.create_sheet('Soldes aux Comptes')
        cols = ['Adhérent', 'Nom du Compte', 'Titre', 'Code ISIN',
                'Solde Disponible', 'Vente en Attente', 'Gelé', 'Nanti', 'Solde Indicatif']
        widths = [12, 38, 42, 16, 18, 18, 10, 10, 18]
        for c, (h, w) in enumerate(zip(cols, widths), 1):
            set_header(ws.cell(row=1, column=c, value=h))
            ws.column_dimensions[get_column_letter(c)].width = w
        ws.row_dimensions[1].height = 30
        ws.freeze_panes = 'A2'
        row = 2
        for acc in soldes:
            for i, l in enumerate(acc.get('lines', [])):
                vals = [acc['adherent'], acc['name'], l['titre'], l['isin'],
                        l['solde_disponible'], l['vente_en_attente'],
                        l['gele'], l['nanti'], l['solde_indicatif']]
                for c, v in enumerate(vals, 1):
                    set_cell(ws.cell(row=row, column=c, value=v), i, c >= 5)
                row += 1

    # ── Valorisation Portefeuille ──────────────────────────────────────
    valorisations = data.get('valorisation', [])
    if valorisations:
        ws = wb.create_sheet('Valorisation Portefeuille')
        cols = ['N° Compte', 'Nom du Compte', 'Titre',
                'Cours Référence', 'Balance', 'Valorisation']
        widths = [12, 38, 42, 16, 14, 22]
        for c, (h, w) in enumerate(zip(cols, widths), 1):
            set_header(ws.cell(row=1, column=c, value=h))
            ws.column_dimensions[get_column_letter(c)].width = w
        ws.row_dimensions[1].height = 30
        ws.freeze_panes = 'A2'
        row = 2
        for acc in valorisations:
            for i, l in enumerate(acc.get('lines', [])):
                vals = [acc['adherent'], acc['name'], l['titre'],
                        l['cours_reference'], l['balance'], l['valorisation']]
                for c, v in enumerate(vals, 1):
                    set_cell(ws.cell(row=row, column=c, value=v), i, c >= 4)
                row += 1
            # Total
            for c in range(1, 7):
                cell = ws.cell(row=row, column=c)
                cell.fill = total_fill
                cell.border = border
                cell.font = Font(bold=True)
            ws.cell(row=row, column=4, value='TOTAL').alignment = Alignment(horizontal='right')
            ws.cell(row=row, column=5,
                    value=sum(l['balance'] for l in acc.get('lines', []))).number_format = num_fmt
            ws.cell(row=row, column=6,
                    value=acc['total_valorisation']).number_format = num_fmt
            row += 2

    # ── Réconciliation ─────────────────────────────────────────────────
    if soldes and valorisations:
        ws = wb.create_sheet('Réconciliation')
        cols = ['Adhérent', 'Nom du Compte', 'Titre',
                'Balance DCBR', 'Balance BTCC', 'Écart', 'Statut']
        widths = [12, 38, 42, 16, 16, 14, 12]
        for c, (h, w) in enumerate(zip(cols, widths), 1):
            set_header(ws.cell(row=1, column=c, value=h))
            ws.column_dimensions[get_column_letter(c)].width = w
        ws.row_dimensions[1].height = 30
        ws.freeze_panes = 'A2'

        # Index soldes by adherent + titre prefix
        soldes_idx = {}
        for acc in soldes:
            adh = acc['adherent']
            soldes_idx[adh] = {}
            for l in acc.get('lines', []):
                key = l['titre'].lower().strip()[:12]
                soldes_idx[adh][key] = l['solde_indicatif']

        row = 2
        for acc in valorisations:
            adh = acc['adherent']
            for l in acc.get('lines', []):
                titre_key = l['titre'].lower().strip()[:12]
                balance_dcbr = None
                if adh in soldes_idx:
                    for k, v in soldes_idx[adh].items():
                        if k[:8] == titre_key[:8]:
                            balance_dcbr = v
                            break

                balance_btcc = l['balance']
                if balance_dcbr is not None:
                    ecart = balance_btcc - balance_dcbr
                    statut = 'OK' if ecart == 0 else 'ÉCART'
                else:
                    ecart = None
                    statut = 'N/A'

                fill = (ok_fill if statut == 'OK' else
                        err_fill if statut == 'ÉCART' else PatternFill())
                vals = [adh, acc['name'], l['titre'],
                        balance_dcbr, balance_btcc, ecart, statut]
                for c, v in enumerate(vals, 1):
                    cell = ws.cell(row=row, column=c, value=v)
                    cell.fill = fill
                    cell.border = border
                    cell.alignment = Alignment(
                        vertical='center',
                        horizontal='right' if 4 <= c <= 6 else 'center' if c == 7 else 'left')
                    if c in (4, 5, 6):
                        cell.number_format = num_fmt
                    if c == 7:
                        cell.font = Font(bold=True,
                                         color=('276749' if statut == 'OK' else
                                                'C53030' if statut == 'ÉCART' else '4A5568'))
                row += 1

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='releves_dcbr_btcc.xlsx'
    )


if __name__ == '__main__':
    app.run(debug=True, port=5000)
