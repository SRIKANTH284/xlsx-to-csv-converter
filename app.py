import os
from flask import Flask, render_template_string, request, send_file
import pandas as pd

app = Flask(__name__)

def convert_xlsx_to_csv(input_file, output_file):
    try:
        df = pd.read_excel(input_file, engine='openpyxl')
        df.to_csv(output_file, index=False)
        print(f"Conversion successful. CSV file saved as '{output_file}'.")
    except Exception as e:
        print(f"An error occurred: {e}")

index_html = '''
<!DOCTYPE html>
<html>
<head>
    <title>XLSX to CSV Converter</title>
</head>
<body>
    <h1>XLSX to CSV Converter</h1>
    {% if error %}
        <p style="color: red;">{{ error }}</p>
    {% endif %}
    {% if message %}
        <p style="color: green;">{{ message }}</p>
        <p><a href="{{ url_for('download_csv') }}">Download CSV</a></p>
    {% else %}
        <form method="post" action="/" enctype="multipart/form-data">
            <input type="file" name="file" accept=".xlsx">
            <input type="submit" value="Upload">
        </form>
    {% endif %}
</body>
</html>
'''

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return render_template_string(index_html, error="No file selected")
        
        file = request.files['file']
        if file.filename == '':
            return render_template_string(index_html, error="No file selected")
        
        if file and file.filename.lower().endswith('.xlsx'):
            # Save the uploaded file temporarily
            temp_path = 'temp.xlsx'
            file.save(temp_path)

            # Convert XLSX to CSV
            output_file_path = 'output.csv'
            convert_xlsx_to_csv(temp_path, output_file_path)

            # Remove the temporary XLSX file
            os.remove(temp_path)

            return render_template_string(index_html, message="Conversion successful. CSV file ready for download.")
        else:
            return render_template_string(index_html, error="Invalid file format. Only XLSX files are supported.")
    
    return render_template_string(index_html)

@app.route('/download')
def download_csv():
    # Provide the link to download the converted CSV file
    return send_file('output.csv', as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
