import os
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from io import BytesIO
from fpdf import FPDF # 🚨 Corrected import for FPDF
from werkzeug.utils import secure_filename
import zipfile 

# --- FLASK CONFIGURATION ---
UPLOAD_FOLDER = 'uploads' 
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
SOURCE_COL = 'source' # Define the key column name once for consistency

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.secret_key = 'super_secret_key_change_me' 

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# --- UTILITY & CORE LOGIC FUNCTIONS ---

def allowed_file(filename):
    """Checks if the file extension is allowed."""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def check_exactly_one_mismatch_or_missing(s1, s2):
    """
    Checks for exactly one character difference (substitution) 
    or one character length difference (insertion/deletion).
    """
    s1 = str(s1).strip()
    s2 = str(s2).strip()

    len1 = len(s1)
    len2 = len(s2)

    if s1 == s2:
        return False

    # 1. Check for exactly one substitution (same length)
    if len1 == len2:
        mismatches = 0
        for i in range(len1):
            if s1[i] != s2[i]:
                mismatches += 1
        return mismatches == 1

    # 2. Check for exactly one insertion/deletion (length difference of 1)
    if abs(len1 - len2) == 1:
        if len1 < len2:
            s1, s2 = s2, s1

        for k in range(len(s1)):
            temp_s1 = s1[:k] + s1[k+1:]
            if temp_s1 == s2:
                return True
        return False
        
    return False

def process_excel_data(file_stream):
    """
    Reads Excel, filters for near-mismatches, and returns the filtered DataFrame
    and the full, unfiltered DataFrame for total count calculation.
    """
    try:
        # Re-seek the stream to the beginning to ensure pandas reads it correctly
        file_stream.seek(0) 
        full_df = pd.read_excel(file_stream, sheet_name='Sheet1', header=0, engine='openpyxl')
        
        # --- Column Definitions ---
        ANPR_COL = 'ANPR Plate Number' 
        TC_COL = 'TC_PLATE Number'
        
        # --- Validation ---
        if ANPR_COL not in full_df.columns or TC_COL not in full_df.columns or SOURCE_COL not in full_df.columns:
              raise ValueError(f"Excel missing required columns: **{ANPR_COL}**, **{TC_COL}**, or **{SOURCE_COL}**.")

        # --- Data Cleaning ---
        full_df[ANPR_COL] = full_df[ANPR_COL].astype(str).str.strip().str.upper() 
        full_df[TC_COL] = full_df[TC_COL].astype(str).str.strip().str.upper()
        full_df[SOURCE_COL] = full_df[SOURCE_COL].astype(str).str.strip()
        
        # --- Filtering Logic ---
        filter_mask = full_df.apply(
            lambda row: check_exactly_one_mismatch_or_missing(
                row[ANPR_COL], 
                row[TC_COL]
            ), 
            axis=1
        )
        
        mismatch_df = full_df[filter_mask].copy() # Use .copy() to prevent SettingWithCopyWarning
        
        return mismatch_df, len(mismatch_df), full_df
    
    except zipfile.BadZipFile as e:
        return None, f"File Error: The uploaded file is corrupted or not a valid Excel (.xlsx) file. Details: {e}", None
    except ValueError as e:
        return None, str(e), None
    except Exception as e:
        print(f"Error during processing: {e}")
        return None, f"An unexpected error occurred during Excel reading: {e}", None

# --- PDF GENERATION LOGIC ---

IMAGE_HEIGHT_MM = 30 
COL_WIDTHS = [25, 30, 25, 55, 60]
COLUMNS_TO_KEEP = ['ANPR Sequence', 'ANPR Plate Number', 'TC_PLATE Number', 'image_path', 'vehicle_image']
SEQUENCE_COL = 'ANPR Sequence'

class ANPR_Report_PDF(FPDF):
    """Custom FPDF class for the ANPR Report layout, including summary data."""
    
    def __init__(self, orientation='P', unit='mm', format='A4', report_title='ANPR Report', total_count=0, mismatch_count=0, zero_seq_count=0):
        super().__init__(orientation, unit, format)
        self.report_title = report_title
        self.total_count = total_count
        self.mismatch_count = mismatch_count
        self.zero_seq_count = zero_seq_count

    def header(self):
        self.set_font('Arial', 'B', 15)
        title_text = f'Near-Mismatch Report: {self.report_title}'
        self.cell(0, 10, title_text, 0, 1, 'C')
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

    def print_header_row(self):
        self.set_fill_color(200, 220, 255) 
        self.set_font('Arial', 'B', 10)
        
        headers = ['Sequence', 'ANPR Plate', 'TC Plate', 'Plate Image', 'Vehicle Image']
        for i, header in enumerate(headers):
            self.cell(COL_WIDTHS[i], 10, header, 1, 0, 'C', 1)
        self.ln()

    def add_data_row(self, data):
        # ... (add_data_row logic: drawing cells and images)
        global IMAGE_HEIGHT_MM, COL_WIDTHS

        if self.get_y() + IMAGE_HEIGHT_MM + 5 > self.page_break_trigger:
            self.add_page()
            self.print_header_row()

        x_start = self.get_x()
        y_start = self.get_y()
        current_x = x_start

        self.set_font('Arial', '', 10)
        
        text_data = [str(data.get(col, '')) for col in COLUMNS_TO_KEEP[:3]]
        
        for i, text in enumerate(text_data):
            self.set_xy(current_x, y_start)
            self.multi_cell(COL_WIDTHS[i], IMAGE_HEIGHT_MM, text, 1, 'C', 0, 0) 
            current_x += COL_WIDTHS[i]

        def insert_image_cell(path_key, col_index):
            nonlocal current_x, y_start
            image_path = data.get(path_key)
            col_width = COL_WIDTHS[col_index]

            self.set_xy(current_x, y_start) 
            self.rect(current_x, y_start, col_width, IMAGE_HEIGHT_MM)
            
            # Simple image logic (assuming image is locally accessible via path)
            try:
                if image_path and os.path.exists(image_path):
                    self.image(image_path, x=current_x+2, y=y_start+2, w=col_width-4, h=IMAGE_HEIGHT_MM-4) 
                else:
                    self.set_xy(current_x + 2, y_start + IMAGE_HEIGHT_MM / 2 - 2)
                    self.set_font('Arial', 'I', 8)
                    self.cell(col_width - 4, 4, 'Image Not Found', 0, 0, 'C')
                    self.set_font('Arial', '', 10) 
            except Exception:
                 self.set_xy(current_x + 2, y_start + IMAGE_HEIGHT_MM / 2 - 2)
                 self.set_font('Arial', 'I', 8)
                 self.cell(col_width - 4, 4, 'Image Load Error', 0, 0, 'C')
                 self.set_font('Arial', '', 10)

            current_x += col_width

        insert_image_cell('image_path', 3)  
        insert_image_cell('vehicle_image', 4) 

        self.set_xy(x_start, y_start + IMAGE_HEIGHT_MM)
    
    def print_summary(self):
        """Prints the summary statistics at the end of the report."""
        self.ln(10)
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, '--- Report Summary ---', 0, 1, 'L')
        
        self.set_font('Arial', '', 10)
        
        # Line 1: Total Vehicles
        self.cell(100, 7, 'Total Vehicles in this Source/Lane:', 1, 0, 'L')
        self.cell(30, 7, str(self.total_count), 1, 1, 'C')
        
        # Line 2: Vehicles with Zero/Unidentified Sequence
        self.cell(100, 7, 'Vehicles with Zero/Unidentified Sequence:', 1, 0, 'L')
        self.cell(30, 7, str(self.zero_seq_count), 1, 1, 'C') 
        
        # Line 3: Near Mismatches
        self.cell(100, 7, 'Near-Mismatched Vehicles:', 1, 0, 'L')
        self.cell(30, 7, str(self.mismatch_count), 1, 1, 'C')
        
        # Line 4: Mismatch Percentage
        if self.total_count > 0:
            percentage = (self.mismatch_count / self.total_count) * 100
            self.set_font('Arial', 'B', 10)
            self.cell(100, 7, 'Mismatch Percentage:', 1, 0, 'L')
            self.cell(30, 7, f'{percentage:.2f}%', 1, 1, 'C')
        self.set_font('Arial', '', 10)


def create_pdf_report(df_mismatch, df_full_source, source_name):
    """Generates a single PDF report with summary statistics."""
    
    df_report = df_mismatch.reindex(columns=COLUMNS_TO_KEEP, fill_value='')

    # Calculate summary counts for this specific source
    total_vehicles = len(df_full_source)
    mismatch_vehicles = len(df_mismatch)
    
    # Calculate count for vehicles with zero/unidentified sequence
    # Filter for 0 (as string/int), NaN, or empty string
    zero_seq_count = df_full_source[
        (df_full_source[SEQUENCE_COL].astype(str).str.strip() == '0') | 
        (df_full_source[SEQUENCE_COL].isnull()) | 
        (df_full_source[SEQUENCE_COL].astype(str).str.strip() == 'unidentified')
    ].shape[0]

    
    # Initialize PDF object, passing the counts
    pdf = ANPR_Report_PDF(
        'P', 'mm', 'A4', 
        report_title=source_name,
        total_count=total_vehicles,
        mismatch_count=mismatch_vehicles,
        zero_seq_count=zero_seq_count
    )
    pdf.set_auto_page_break(False, margin=15)
    pdf.add_page()

    pdf.print_header_row()
    
    # Process rows
    for index, row in df_report.iterrows():
        row_dict = row.to_dict()
        row_dict['image_path'] = str(row_dict.get('image_path', ''))
        row_dict['vehicle_image'] = str(row_dict.get('vehicle_image', ''))

        pdf.add_data_row(row_dict)

    # Print the summary
    if pdf.get_y() + 50 > pdf.page_break_trigger:
         pdf.add_page()
         
    pdf.print_summary()

    return pdf.output(dest='S')


# --- FLASK ROUTES ---

@app.route('/')
def index():
    """Renders the main upload form."""
    # You must have an 'index.html' template in a 'templates' folder.
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handles file upload, processing, multiple PDF generation, and ZIP download."""
    
    if 'excel_file' not in request.files:
        flash('No file part in the request.', 'error')
        return redirect(url_for('index'))
    
    file = request.files['excel_file']
    
    if file.filename == '':
        flash('No selected file.', 'error')
        return redirect(url_for('index'))

    if file and allowed_file(file.filename):
        try:
            # 1. Process data: returns mismatch_df, count (or error), and full_df
            mismatch_df, count, full_df = process_excel_data(file.stream)

            if isinstance(count, str): 
                flash(f'Processing Error: {count}', 'error')
                return redirect(url_for('index'))

            # 2. Check if mismatches were found
            if mismatch_df.empty:
                flash('✅ File processed successfully! No near-mismatches found.', 'success')
                return redirect(url_for('index'))
            
            # 3. Group filtered and full dataframes by 'source'
            mismatch_groups = mismatch_df.groupby(SOURCE_COL)
            full_groups = full_df.groupby(SOURCE_COL)
            
            # Create an in-memory ZIP file container
            zip_buffer = BytesIO()
            total_pdf_count = 0
            
            with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED, False) as zf:
                
                for source_name, mismatch_group_df in mismatch_groups:
                    
                    # Get the corresponding full data for the current source
                    # This is used for the "Total Vehicles" count in the summary
                    full_group_df = full_groups.get_group(source_name) 
                    
                    if not source_name or str(source_name).strip().lower() == 'nan':
                        source_name = "UNSPECIFIED_SOURCE"

                    # Generate the PDF for this group
                    pdf_bytes = create_pdf_report(mismatch_group_df, full_group_df, source_name)
                    
                    # Add the PDF to the ZIP file
                    safe_source_name = secure_filename(source_name) 
                    pdf_filename = f"ANPR_Report_{safe_source_name}.pdf"
                    
                    zf.writestr(pdf_filename, pdf_bytes)
                    total_pdf_count += 1
            
            zip_buffer.seek(0)
            
            # 4. Return the ZIP file for download
            base_name = secure_filename(file.filename.rsplit('.', 1)[0])
            output_zip_filename = f"ANPR_Reports_Grouped_By_Source_{base_name}.zip"

            flash(f'✅ Found **{count}** total near-mismatches, grouped into **{total_pdf_count}** reports. Downloading ZIP file.', 'success')
            
            return send_file(
                zip_buffer,
                mimetype='application/zip',
                as_attachment=True,
                download_name=output_zip_filename
            )

        except Exception as e:
            flash(f'An unexpected error occurred during processing or report generation: {e}', 'error')
            return redirect(url_for('index'))
    
    else:
        flash('Invalid file type. Only .xlsx and .xls are allowed.', 'error')
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)