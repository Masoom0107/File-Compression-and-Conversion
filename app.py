from flask import Flask, render_template, request, send_from_directory, send_file , abort ,flash, redirect,url_for 
import threading
import logging
import zipfile
import tempfile
import os
from docx import Document
import io
import shutil
from moviepy.editor import VideoFileClip
from werkzeug.utils import secure_filename
import sys
from PIL import Image
from pptx import Presentation
import openpyxl
import time
import aspose.pdf as ap
from pdf2docx import parse
import comtypes.client
from openpyxl import load_workbook
from reportlab.lib.pagesizes import landscape, letter
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch 
import pandas as pd
import tabula
from pdfminer.high_level import extract_text  #pdf to text
from fpdf import FPDF  #text to pdf
from moviepy.editor import VideoFileClip #mp4 to mkv and mp4 to gif

app = Flask(__name__)
app.static_folder = 'static'
app.config['UPLOAD_FOLDER'] = 'uploads'
# Configuration
app.config['UPLOAD_FOLDER'] = 'uploads'
app.secret_key = 'your_secret_key_here'


class FileCompression:
    def __init__(self, file_type):
        self.file_type = file_type
        self.new_filename = ""

    # Define the compression method for each file type
    def compress(self, file_path):
        if self.file_type == "video":
            # Implement video compression logic here
            pass
        elif self.file_type == "pdf":
            # Implement PDF compression logic here
            pass
        elif self.file_type == "image":
            # Implement image compression logic here
            pass
        elif self.file_type == "zip":
            # Implement zip compression logic here
            pass
        # Add more conditions for other file types
        elif self.file_type == "excel":
            return self.compress_excel(file_path)

    def compress_excel(self, file_path):
        if file_path.endswith(".xlsx"):
            try:
                workbook = openpyxl.load_workbook(file_path)
                output_file_path = file_path.replace(".xlsx", "_compressed.xlsx")
                workbook.save(output_file_path)
                return output_file_path
            except Exception as e:
                print("Error during Excel compression:", e)
                return None
        else:
            return "Invalid file format. Please provide an Excel (.xlsx) file."


# Create instances for different compression types
video_compression_tool = FileCompression("video")
pdf_compression_tool = FileCompression("pdf")
image_compression_tool = FileCompression("image")
zip_compression_tool = FileCompression("zip")
excel_compression_tool = FileCompression("excel")

UPLOAD_FOLDER = 'uploads'

@app.route('/pdf_to_word', methods=['GET', 'POST'])
def pdf_to_word():
    if request.method == 'POST':
        if 'pdf' not in request.files:
            return render_template('pdf_to_word.html', error="No PDF file selected.")

        pdf_file = request.files['pdf']

        if pdf_file.filename == '':
            return render_template('pdf_to_word.html', error="No PDF file selected.")

        # Check if the file has a PDF extension
        if not pdf_file.filename.lower().endswith('.pdf'):
            return render_template('pdf_to_word.html', error="Invalid file format. Please select a PDF file.")

        # Save the PDF file to a temporary location
        pdf_filename = secure_filename(pdf_file.filename)
        pdf_path = os.path.join(UPLOAD_FOLDER, pdf_filename)
        pdf_file.save(pdf_path)

        # Get the output path for the Word file
        output_filename = 'converted.docx'
        output_path = os.path.join(UPLOAD_FOLDER, output_filename)

        # Convert PDF to Word
        try:
            parse(pdf_path, output_path, start=0, end=None)
        except Exception as e:
            return render_template('pdf_to_word.html', error=f"Error: {str(e)}")

        # Provide a link to download the converted Word file
        download_link = f'<a href="/download/{output_filename}">Download Converted File</a>'

        return render_template('pdf_to_word.html', result=f"File converted successfully. {download_link}")

    return render_template('pdf_to_word.html')

@app.route('/word_to_pdf', methods=['GET', 'POST'])
def word_to_pdf():
    if request.method == 'POST':
        file = request.files['word']
        if file.filename != '':
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
            file.save(file_path)

            output_path = os.path.splitext(file_path)[0] + '.pdf'
            convert_word_to_pdf(file_path, output_path)
            return send_file(output_path, as_attachment=True)

    return render_template('word_to_pdf.html')

def convert_word_to_pdf(input_path, output_path):
    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = False
    docx_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)
    in_file = word.Documents.Open(docx_path)
    in_file.SaveAs(output_path, FileFormat=17)
    in_file.Close()
    word.Quit()

def convert_to_pdf(input_path, output_path):
    # Convert Word to PDF using pdf2docx library
    cv = Converter(input_path)
    cv.convert(output_path)
    cv.close()

ALLOWED_EXTENSIONS_PDF = {'pdf'}

def allowed_file_pdf(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS_PDF

ALLOWED_EXTENSIONS_PDF = {'pdf'}

def allowed_file_pdf(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS_PDF

@app.route('/pdf_to_excel', methods=['GET', 'POST'])
def pdf_to_excel():
    if request.method == 'POST':
        # Check if the post request has the file part
        if 'pdf' not in request.files:
            flash('No file part')
            return redirect(request.url)
        
        pdf_file = request.files['pdf']
        
        # If user does not select file, browser also submit an empty part without filename
        if pdf_file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        
        # Check if the file extension is allowed
        if pdf_file and allowed_file_pdf(pdf_file.filename):
            # Define paths
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(pdf_file.filename))
            output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'output.xlsx')
            
            # Save the uploaded PDF file
            pdf_file.save(file_path)
            
            # Convert PDF to CSV
            tabula.convert_into(file_path, output_path, output_format="csv", pages='1')
            
            # Convert CSV to Excel
            data = pd.read_csv(output_path)
            data.to_excel(output_path, index=False)
            
            # Provide feedback to the user
            output_filename = 'output.xlsx'  # Define the output filename
            download_link = f'<a href="{url_for("download_file", filename=output_filename)}">Download Converted File</a>'
            return render_template('pdf_to_excel.html', result='File converted successfully.', download_link=download_link)
                
        else:
            flash('Only PDF files are allowed')
            return redirect(request.url)
        
    return render_template('pdf_to_excel.html')

ALLOWED_EXTENSIONS_PDF = {'pdf'}

def allowed_file_pdf(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS_PDF

@app.route('/pdf_to_csv', methods=['GET', 'POST'])
def pdf_to_csv():
    if request.method == 'POST':
        # Check if the post request has the file part
        if 'pdf' not in request.files:
            flash('No file part')
            return redirect(request.url)
        
        pdf_file = request.files['pdf']
        
        # If user does not select file, browser also submit an empty part without filename
        if pdf_file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        
        # Check if the file extension is allowed
        if pdf_file and allowed_file_pdf(pdf_file.filename):
            # Define paths
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(pdf_file.filename))
            output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'output.csv')
            
            # Save the uploaded PDF file
            pdf_file.save(file_path)
            
            # Convert PDF to CSV
            tabula.convert_into(file_path, output_path, pages="all", output_format="csv")
            
            # Provide feedback to the user
            output_filename = 'output.csv'  # Define the output filename
            download_link = f'<a href="{url_for("download_file", filename=output_filename)}">Download Converted File</a>'
            return render_template('pdf_to_csv.html', result='File converted successfully.', download_link=download_link)
        
        else:
            flash('Only PDF files are allowed')
            return redirect(request.url)
        
    return render_template('pdf_to_csv.html')

@app.route('/pdf_to_text', methods=['GET', 'POST'])
def pdf_to_text():
    if request.method == 'POST':
        # Check if the post request has the file part
        if 'pdf' not in request.files:
            flash('No file part')
            return redirect(request.url)
        
        pdf_file = request.files['pdf']
        
        # If user does not select file, browser also submit an empty part without filename
        if pdf_file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        
        # Check if the file extension is allowed
        if pdf_file and allowed_file_pdf(pdf_file.filename):
            # Define paths
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(pdf_file.filename))
            output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'output.txt')
            
            # Save the uploaded PDF file
            pdf_file.save(file_path)
            
            # Convert PDF to text
            text = extract_text(file_path)
            
            # Write text to output file
            with open(output_path, "w") as output_file:
                output_file.write(text)
            
            # Provide feedback to the user
            output_filename = 'output.txt'  # Define the output filename
            download_link = f'<a href="{url_for("download_file", filename=output_filename)}">Download Converted File</a>'
            return render_template('pdf_to_text.html', result='File converted successfully.', download_link=download_link)
        
        else:
            flash('Only PDF files are allowed')
            return redirect(request.url)
        
    return render_template('pdf_to_text.html')

# Routes for Excel conversions
ALLOWED_EXTENSIONS_EXCEL = {'xlsx'}

def allowed_file_excel(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS_EXCEL

@app.route('/excel_to_pdf', methods=['GET', 'POST'])
def excel_to_pdf():
    if request.method == 'POST':
        # Check if the post request has the file part
        if 'excel' not in request.files:
            flash('No file part')
            return redirect(request.url)
        
        excel_file = request.files['excel']
        
        # If user does not select file, browser also submit an empty part without filename
        if excel_file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        
        # Check if the file extension is allowed
        if excel_file and allowed_file_excel(excel_file.filename):
            # Define paths
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(excel_file.filename))
            output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'output.pdf')
            
            # Save the uploaded Excel file
            excel_file.save(file_path)
            
            # Convert Excel to PDF
            workbook = load_workbook(file_path)
            worksheet = workbook.active
            max_row = worksheet.max_row
            max_column = worksheet.max_column
            
            # Create PDF canvas
            c = canvas.Canvas(output_path, pagesize=landscape(letter))
            top_margin = 3 * inch
            left_margin = 0.5 * inch
            bottom_margin = 0.5 * inch
            right_margin = 0.5 * inch
            cell_width = (11 * inch - left_margin - right_margin) / max_column
            cell_height = (8.5 * inch - top_margin - bottom_margin) / max_row
            
            # Write Excel content to PDF
            for row in range(1, max_row + 1):
                for column in range(1, max_column + 1):
                    cell = worksheet.cell(row=row, column=column)
                    text = str(cell.value)
                    x = left_margin + (column - 1) * cell_width
                    y = 11 * inch - (top_margin + row * cell_height)
                    c.drawString(x, y, text)
            
            # Save PDF file
            c.save()
            
            # Provide feedback to the user
            output_filename = 'output.pdf'  # Define the output filename
            download_link = f'<a href="{url_for("download_file", filename=output_filename)}">Download Converted File</a>'
            return render_template('excel_to_pdf.html', result='File converted successfully.', download_link=download_link)

        
        else:
            flash('Only Excel files (xlsx) are allowed')
            return redirect(request.url)
        

    return render_template('excel_to_pdf.html')

# Routes for Text conversions
ALLOWED_EXTENSIONS_TXT = {'txt'}

def allowed_file_txt(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS_TXT

@app.route('/text_to_pdf', methods=['GET', 'POST'])
def txt_to_pdf():
    if request.method == 'POST':
        # Check if the post request has the file part
        if 'txt' not in request.files:
            flash('No file part')
            return redirect(request.url)
        
        txt_file = request.files['txt']
        
        # If user does not select file, browser also submit an empty part without filename
        if txt_file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        
        # Check if the file extension is allowed
        if txt_file and allowed_file_txt(txt_file.filename):
            # Define paths
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(txt_file.filename))
            output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'output.pdf')
            
            # Save the uploaded TXT file
            txt_file.save(file_path)
            
            # Convert TXT to PDF
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=15)
            
            # Set left and right margins
            left_margin = 10
            right_margin = 10
            
            # Read text from TXT file and add to PDF
            with open(file_path, "r", encoding='utf-8', errors='ignore') as txt_file:
                for line in txt_file:
                    # Calculate the available width for the cell
                    available_width = 190 - left_margin - right_margin
                    # Calculate the width of the text
                    text_width = pdf.get_string_width(line)
                    # If the text width exceeds the available width, adjust the width of the cell
                    if text_width > available_width:
                        pdf.multi_cell(available_width, 10, txt=line.strip(), align='L')
                    else:
                        pdf.cell(0, 10, txt=line.strip(), ln=True, align='L')            
            
            # Save PDF file
            pdf.output(output_path)
            
            # Provide feedback to the user
            output_filename = 'output.pdf'  # Define the output filename
            download_link = f'<a href="{url_for("download_file", filename=output_filename)}">Download Converted File</a>'
            return render_template('text_to_pdf.html', result='File converted successfully.', download_link=download_link)
        
        else:
            flash('Only TXT files are allowed')
            return redirect(request.url)
        
    return render_template('text_to_pdf.html')

# Routes for Video conversions
ALLOWED_EXTENSIONS_MP4 = {'mp4'}
def allowed_file_mp4(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS_MP4

@app.route('/mp4_to_mkv', methods=['GET', 'POST'])
def mp4_to_mkv():
    if request.method == 'POST':
        if 'mp4' not in request.files:
            flash('No file part')
            return redirect(request.url)
        
        mp4_file = request.files['mp4']
        
        if mp4_file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        
        if mp4_file and allowed_file_mp4(mp4_file.filename):
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'input.mp4')
            output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'output.mkv')
            print("File path:", file_path)
            print("Output path:", output_path)
            mp4_file.save(file_path)
            print("File saved successfully")
            
            clip = VideoFileClip(file_path)
            clip.write_videofile(output_path, codec="libx264", preset="medium")
            print("Conversion completed")
            
            download_link = f'<a href="{url_for("download_file", filename="output.mkv")}">Download Converted File</a>'
            return render_template('mp4_to_mkv.html', result='File converted successfully.', download_link=download_link)
        else:
            flash('Only MP4 files are allowed')
            return redirect(request.url)
        
    return render_template('mp4_to_mkv.html')

@app.route('/mp4_to_gif')
def mp4_to_gif():
    # Implement MP4 to GIF conversion logic here
    return render_template('mp4_to_gif', result="MP4 to GIF Conversion")


@app.route('/converter', methods=['GET','POST'])
def converter():
    return render_template('converter.html')

@app.route('/compressor', methods=['GET','POST'])
def compressor():
    return render_template('compressor.html')

@app.route('/selection', methods=['GET', 'POST'])
def selection():
    return render_template('selection_page.html')



@app.route('/compress_excel', methods=['GET', 'POST'])
def compress_excel():
    if request.method == 'POST':
        if 'excel' in request.files:
            excel_file = request.files['excel']
            if excel_file.filename != '':
                excel_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(excel_file.filename))
                excel_file.save(excel_path)

                compressed_filename = excel_compression_tool.compress(excel_path)

                if compressed_filename:
                    # Compression was successful, so return a link to download the compressed file
                    return render_template('excel_compression.html',
                                           result=f"File compressed successfully. <a href='/download/{os.path.basename(compressed_filename)}'>Download Compressed File</a>")
                else:
                    # Compression failed
                    return render_template('excel_compression.html', result="Compression failed")

    return render_template('excel_compression.html')


class ImageCompression:
    def __init__(self):
        self.image_path = ""
        self.compression_quality = "best"
        self.new_filename = ""

    def compress_image(self):
        image_path = self.image_path
        quality = self.compression_quality
        quality_mapping = {'good': 60, 'better': 30, 'best': 9}
        compression_quality = quality_mapping.get(quality, 60)
        new_filename = self.new_filename + '.jpg'

        try:
            img = Image.open(image_path)
            compressed_img = img.copy()
            compressed_img.save(new_filename, format='JPEG', quality=compression_quality)
            return new_filename
        except Exception as e:
            print("Error during compression:", e)
            return None


image_compression_tool = ImageCompression()


@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')


# Function to compress a PDF file
def compress_pdf_file(pdf_path):
    try:
        pdf_document = ap.Document(pdf_path)
        compressed_path = pdf_path.replace('.pdf', '_compressed.pdf')

        pdf_optimization = ap.optimization.OptimizationOptions()
        pdf_optimization.image_compression_options.compress_images = True
        pdf_optimization.image_compression_options.image_quality = 50

        pdf_document.optimize_resources(pdf_optimization)
        pdf_document.save(compressed_path)

        return compressed_path
    except Exception as e:
        print("Error during PDF compression:", e)
        return None


@app.route('/compress_pdf', methods=['GET', 'POST'])
def compress_pdf():
    if request.method == 'POST':
        pdf_file = request.files['pdf']
        if pdf_file.filename != '':
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(pdf_file.filename))
            pdf_file.save(pdf_path)

            compressed_filename = compress_pdf_file(pdf_path)

            if compressed_filename:
                # Compression was successful, so return a link to download the compressed file
                return render_template('pdf_compression.html',
                                       result="File compressed successfully. <a href='/download/" + os.path.basename(
                                           compressed_filename) + "'>Download Compressed File</a>")
            else:
                # Compression failed
                return render_template('pdf_compression.html', result="Compression failed")

    return render_template('pdf_compression.html')


@app.route('/compress_images', methods=['GET', 'POST'])
def compress_images():
    if request.method == 'POST':
        if 'image' in request.files:
            image_file = request.files['image']
            if image_file.filename != '':
                image_path = os.path.join(app.config['UPLOAD_FOLDER'], image_file.filename)
                image_file.save(image_path)
                image_compression_tool.image_path = image_path
                image_compression_tool.compression_quality = request.form.get('compression_quality')
                image_compression_tool.new_filename = request.form.get('new_filename')

                compressed_filename = image_compression_tool.compress_image()
                if compressed_filename:
                    compressed_path = os.path.join(app.config['UPLOAD_FOLDER'], compressed_filename)
                    shutil.move(compressed_filename, compressed_path)
                    
                    # Generate a download link for the compressed image
                    download_link = f'<a href="/download/{os.path.basename(compressed_path)}">Download Compressed Image</a>'
                    
                    # Render the template with the download link
                    return render_template('image_compression.html', result=download_link)

    return render_template('image_compression.html')




@app.route('/compress_videos', methods=['GET', 'POST'])
def compress_videos():
    if request.method == 'POST':
        try:
            file = request.files['video']
            if file and file.filename != '':
                video_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
                file.save(video_path)

                preset = request.form.get('compression_preset')
                compression_percentage = get_compression_percentage(preset)
                new_filename = request.form.get('new_filename') + '_compressed.mp4'

                result = compress_video_with_moviepy(video_path, compression_percentage, new_filename)

                if result:
                    # Move the compressed video to the uploads folder
                    compressed_path = os.path.join(app.config['UPLOAD_FOLDER'],  new_filename)
                    shutil.move(new_filename, compressed_path)

                    # Generate a download link for the compressed video
                    download_link = f'<a href="/download/{os.path.basename(compressed_path)}">Download Compressed Video</a>'

                    return render_template('video_compression.html', result=download_link)
                else:
                    return render_template('video_compression.html', result='Error: Compression Failed', result_color='red')
        except Exception as e:
           return render_template('video_compression.html', result=f'Error: {str(e)}', result_color='red')

    return render_template('video_compression.html')







def get_compression_percentage(preset):
    presets = {'good': 0.8, 'better': 0.6, 'best': 0.4}
    return presets.get(preset, 1.0)


def compress_video_with_moviepy(video_path, compression_percentage, new_filename):
    try:
        clip = VideoFileClip(video_path)
        compressed_clip = clip.subclip().fx(VideoFileClip.resize, compression_percentage)
        compressed_clip.write_videofile(new_filename)
        return True
    except Exception as e:
        print("Error during compression:", e)
        return False


class FileCompression:
    def __init__(self, file_type):
        self.file_type = file_type
        self.new_filename = ""

    # Define the compression method for each file type
    def compress(self, file_path):
        if self.file_type == "zip":
            return self.compress_zip(file_path)
        # Add more conditions for other file types
        elif self.file_type == "excel":
            return self.compress_excel(file_path)

    def compress_zip(self, input_zip_path):
        try:
            # Create a temporary directory to extract the original ZIP file
            temp_dir = tempfile.mkdtemp()

            # Extract the original ZIP file to the temporary directory
            with zipfile.ZipFile(input_zip_path, 'r') as input_zip:
                input_zip.extractall(temp_dir)

            # Compress individual files within the temporary directory
            compressed_folder = os.path.join(temp_dir, 'compressed')
            os.makedirs(compressed_folder, exist_ok=True)
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    shutil.make_archive(os.path.join(compressed_folder, file), 'zip', root, file)

            # Create a new ZIP archive with the compressed files
            output_zip_path = input_zip_path.replace('.zip', '_compressed.zip')
            with zipfile.ZipFile(output_zip_path, 'w', zipfile.ZIP_DEFLATED) as output_zip:
                for root, dirs, files in os.walk(compressed_folder):
                    for file in files:
                        file_path = os.path.join(root, file)
                        archive_path = os.path.relpath(file_path, compressed_folder)
                        output_zip.write(file_path, archive_path)

            # Clean up the temporary directory
            shutil.rmtree(temp_dir)

            return output_zip_path
        except Exception as e:
            print("Error during ZIP compression:", e)
            return None

    def compress_excel(self, file_path):
        if file_path.endswith(".xlsx"):
            try:
                # Implement Excel compression logic here
                # Example: Load the Excel file, make changes, and save it as a compressed file
                # For now, we'll return the same input file as a placeholder
                return file_path
            except Exception as e:
                print("Error during Excel compression:", e)
                return None
        else:
            return "Invalid file format. Please provide an Excel (.xlsx) file."


# Create an instance for ZIP compression
zip_compression_tool = FileCompression("zip")


@app.route('/compress_zip', methods=['GET', 'POST'])
def compress_zip():
    if request.method == 'POST':
        zip_file = request.files['zip']
        if zip_file.filename != '':
            zip_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(zip_file.filename))
            zip_file.save(zip_path)

            compressed_filename = zip_compression_tool.compress(zip_path)

            if compressed_filename:
                # Compression was successful, so return a link to download the compressed file
                return render_template('zip_compression.html', result="File compressed successfully. <a href='/download/" + os.path.basename(compressed_filename) + "'>Download Compressed File</a>")
            else:
                # Compression failed
                return render_template('zip_compression.html', result="Compression failed")

    return render_template('zip_compression.html')


@app.route('/compress_docx', methods=['GET', 'POST'])
def compress_docx():
    if request.method == 'POST':
        if 'docx_file' in request.files:
            docx_file = request.files['docx_file']
            if docx_file.filename != '':
                docx_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(docx_file.filename))
                docx_file.save(docx_path)

                compressed_filename = compress_docx_file(docx_path)

                if compressed_filename:
                    # Compression was successful, so return a link to download the compressed file
                    return render_template('docx_compression.html',
                                           result="File compressed successfully. <a href='/download/" + os.path.basename(
                                               compressed_filename) + "'>Download Compressed File</a>")
                else:
                    # Compression failed
                    return render_template('docx_compression.html', result="Compression failed")

    return render_template('docx_compression.html')


def compress_docx_file(docx_path):
    try:
        doc = Document(docx_path)
        compressed_file_path = docx_path.replace('.docx', '_compressed.docx')

        for image in doc.inline_shapes:
            if hasattr(image, 'image'):
                original_image = image.image.blob
                compressed_image = compress_image(original_image)
                image.image.blob = compressed_image

        doc.save(compressed_file_path)
        return compressed_file_path
    except Exception as e:
        print("Error during DOCX compression:", e)
        return None


def compress_image(image_blob):
    original_image = Image.open(io.BytesIO(image_blob))
    compressed_image = original_image.copy()
    compressed_image.thumbnail((800, 600))
    buffer = io.BytesIO()
    compressed_image.save(buffer, format='PNG')
    return buffer.getvalue()


@app.route('/compress_ppt', methods=['GET', 'POST'])
def compress_ppt():
    if request.method == 'POST':
        ppt_file = request.files['ppt']
        if ppt_file.filename != '':
            ppt_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(ppt_file.filename))
            ppt_file.save(ppt_path)

            compressed_filename = compress_ppt_file(ppt_path)

            if compressed_filename:
                # Compression was successful, so return a link to download the compressed file
                return render_template('ppt_compression.html',
                                       result="File compressed successfully. <a href='/download/" + os.path.basename(
                                           compressed_filename) + "'>Download Compressed File</a>")
            else:
                # Compression failed
                return render_template('ppt_compression.html', result="Compression failed")

    return render_template('ppt_compression.html')


def compress_ppt_file(ppt_path):
    try:
        prs = Presentation(ppt_path)
        compressed_path = ppt_path.replace('.pptx', '_compressed.pptx')

        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "element") and hasattr(shape.element, "get_or_add_image"):
                    image = shape.element.get_or_add_image()
                    image.compression = 60

        prs.save(compressed_path)
        return compressed_path
    except Exception as e:
        print("Error during PPT compression:", e)
        return None


@app.route('/archive_creation', methods=['GET', 'POST'])
def archive_creation():
    if request.method == 'POST':
        file_paths = request.files.getlist('files[]')

        if not file_paths:
            return render_template('archive_creation.html', archive_creation_status="No files selected.")

        # Generate a unique archive name based on the current timestamp
        timestamp = int(time.time())
        zip_file_path = os.path.join(app.config['UPLOAD_FOLDER'], f'archive_{timestamp}.zip')

        try:
            compression = zipfile.ZIP_DEFLATED
            with zipfile.ZipFile(zip_file_path, 'w') as zipf:
                for file in file_paths:
                    file.save(os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename)))
                    base_name = secure_filename(file.filename)
                    zipf.write(os.path.join(app.config['UPLOAD_FOLDER'], base_name), base_name, compress_type=compression)

            # Provide a link to download the created archive
            archive_filename = os.path.basename(zip_file_path)
            return render_template('archive_creation.html',
                                   archive_creation_status=f"Archive created successfully. <a href='/download/{archive_filename}'>Download Created Archive</a>")
        except Exception as e:
            return render_template('archive_creation.html', archive_creation_status=f"Archive creation failed: {str(e)}")

    return render_template('archive_creation.html')


@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join(app.config['UPLOAD_FOLDER'], filename), as_attachment=True)

@app.route('/download/<filename>', methods=['GET'])
def download(filename):
    compressed_dir = os.path.join(app.config['UPLOAD_FOLDER'])
    file_path = os.path.join(compressed_dir, filename)
    
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        abort(404)  # Return a 404 error if the file doesn't exist

if __name__ == '__main__':
    app.run(debug=True)
