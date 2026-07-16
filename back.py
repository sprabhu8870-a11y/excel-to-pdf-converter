import io
import traceback
from flask import Flask, request, send_file, jsonify, render_template
import pandas as pd
from reportlab.lib.pagesizes import letter, A4, legal, landscape, portrait
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

app = Flask(__name__, template_folder='templates')

PAGE_SIZES = {
    'a4': A4,
    'letter': letter,
    'legal': legal
}

def build_pdf_stream(file_stream, filename, orientation, paper_size, show_gridlines, convert_all_sheets, fit_columns):
    """Parses spreadsheet stream dynamically depending on file format and builds an in-memory PDF."""
    # 1. Page Frame Geometry Matrix
    base_size = PAGE_SIZES.get(paper_size.lower(), A4)
    page_dimension = landscape(base_size) if orientation == 'landscape' else portrait(base_size)
        
    page_width, page_height = page_dimension
    margin = 24  
    usable_width = page_width - (margin * 2)

    # 2. Dynamic File Parsing Engine Wrapper
    file_extension = filename.split('.')[-1].lower()
    
    sheets_data = {}
    
    if file_extension == 'csv':
        # CSV files only have one implicit sheet
        df = pd.read_csv(file_stream, header=None, dtype=str)
        sheets_data['CSV_Export'] = df
    elif file_extension == 'xls':
        # Legacy Excel files require 'xlrd' engine
        excel_file = pd.ExcelFile(file_stream, engine='xlrd')
        sheets_to_process = excel_file.sheet_names if convert_all_sheets else [excel_file.sheet_names[0]]
        for sheet in sheets_to_process:
            sheets_data[sheet] = pd.read_excel(excel_file, sheet_name=sheet, header=None, dtype=str)
    else:
        # Standard modern OpenXML workbook format (.xlsx)
        excel_file = pd.ExcelFile(file_stream, engine='openpyxl')
        sheets_to_process = excel_file.sheet_names if convert_all_sheets else [excel_file.sheet_names[0]]
        for sheet in sheets_to_process:
            sheets_data[sheet] = pd.read_excel(excel_file, sheet_name=sheet, header=None, dtype=str)

    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=page_dimension,
        leftMargin=margin,
        rightMargin=margin,
        topMargin=margin,
        bottomMargin=margin
    )
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'SheetTitle',
        parent=styles['Heading2'],
        textColor=colors.HexColor('#059669'),
        spaceAfter=12
    )
    cell_style = ParagraphStyle(
        'TableCell',
        parent=styles['Normal'],
        fontSize=8,
        leading=10,
        textColor=colors.HexColor('#1e293b')
    )
    header_style = ParagraphStyle(
        'TableHeader',
        parent=styles['Normal'],
        fontSize=8,
        leading=10,
        textColor=colors.white,
        fontName='Helvetica-Bold'
    )

    story = []

    for idx, (sheet_name, df) in enumerate(sheets_data.items()):
        df = df.fillna('') 

        # Drop outer padding empty structural rows
        while len(df) > 0 and df.iloc[0].astype(str).str.strip().eq('').all():
            df = df.iloc[1:].reset_index(drop=True)

        if df.empty:
            continue

        if idx > 0:
            from reportlab.platypus import PageBreak
            story.append(PageBreak())

        story.append(Paragraph(f"Sheet: {sheet_name}", title_style))
        
        table_data = []
        raw_matrix = df.values.tolist()
        
        for row_idx, raw_row in enumerate(raw_matrix):
            formatted_row = []
            for cell in raw_row:
                current_style = header_style if row_idx == 0 else cell_style
                clean_text = str(cell).replace('\r', '').replace('\n', '<br/>')
                formatted_row.append(Paragraph(clean_text, current_style))
            table_data.append(formatted_row)

        col_count = len(raw_matrix[0]) if raw_matrix else 0
        
        if col_count > 0:
            if fit_columns:
                col_widths = [usable_width / col_count] * col_count
            else:
                col_widths = [75] * col_count
        else:
            col_widths = None

        reportlab_table = Table(table_data, colWidths=col_widths, repeatRows=1)
        
        t_style = [
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#059669')),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]
        
        if show_gridlines:
            t_style.extend([
                ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#cbd5e1')),
                ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#94a3b8'))
            ])
            
        reportlab_table.setStyle(TableStyle(t_style))
        story.append(reportlab_table)
        story.append(Spacer(1, 15))

    if not story:
        raise ValueError("Workbook target structure possesses zero active rows data.")

    doc.build(story)
    pdf_buffer.seek(0)
    return pdf_buffer

@app.route('/')
def index():
    return render_template('front.html')

@app.route('/api/convert', methods=['POST'])
def convert_excel_to_pdf():
    try:
        if 'excelFile' not in request.files:
            return jsonify({'error': 'Spreadsheet payload missing in structural content container.'}), 400
            
        file = request.files['excelFile']
        if file.filename == '':
            return jsonify({'error': 'Empty filename parameters identified.'}), 400

        orientation = request.form.get('orientation', 'landscape')
        paper_size = request.form.get('paperSize', 'a4')
        show_gridlines = request.form.get('gridlines', 'true') == 'true'
        convert_all_sheets = request.form.get('allSheets', 'false') == 'true'
        fit_columns = request.form.get('scaling', 'fit-columns') == 'fit-columns'

        # Wrap raw bytes block inside an in-memory binary builder stream
        file_bytes = file.read()
        file_stream = io.BytesIO(file_bytes)

        pdf_stream = build_pdf_stream(
            file_stream=file_stream,
            filename=file.filename,
            orientation=orientation,
            paper_size=paper_size,
            show_gridlines=show_gridlines,
            convert_all_sheets=convert_all_sheets,
            fit_columns=fit_columns
        )

        base_name = file.filename.rsplit('.', 1)[0]
        return send_file(
            pdf_stream,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f"{base_name}.pdf"
        )

    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': f"Conversion breakdown: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)