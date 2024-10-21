from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
from data_cleaning import semantic_mapping, save_to_excel_with_highlight

app = Flask(__name__)

@app.route('/')
def home():
    return render_template('chat.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    try:
        # Retrieve the uploaded files
        veeva_file = request.files.get('veeva')
        xponent_file = request.files.get('xponent')

        # Check if both files are uploaded
        if not veeva_file or not xponent_file:
            return jsonify({"message": "Files not received correctly"}), 400

        # Load the data from the files into pandas DataFrames
        veeva_df = pd.read_excel(veeva_file)
        xponent_df = pd.read_excel(xponent_file)

        # Perform semantic mapping
        output_file_path = 'output/cleaned_xponent.xlsx'  # Save as Excel file
        semantic_mapping(veeva_df, xponent_df, output_file_path)  # Pass the output file name

        # Return the top 5 rows for preview
        cleaned_df = pd.read_excel(output_file_path)
        top_5 = cleaned_df.head().to_html(classes='table table-bordered')
        return jsonify({
            "message": "Mapping completed successfully!",
            "top_5": top_5
        })
    
    except Exception as e:
        print(f"Error occurred: {e}")
        return jsonify({"message": f"Error occurred: {e}"}), 500


@app.route('/download', methods=['GET'])
def download_file():
    try:
        output_file_path = 'output/cleaned_xponent.xlsx'  # Adjusted to point to the Excel file
        return send_file(output_file_path, as_attachment=True)
    except Exception as e:
        return jsonify({"message": f"Error occurred: {e}"}), 500

if __name__ == '__main__':
    app.run(debug=True)
