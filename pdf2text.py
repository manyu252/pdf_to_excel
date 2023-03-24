import os
import pandas as pd
import json
from flask import Flask, render_template, request, redirect, url_for, send_file
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024
ALLOWED_EXTENSIONS = {'pdf'}
UPLOAD_FOLDER = os.getcwd() + '/uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

deposit_search_keyword = "DEPOSIT"
withdrawal_search_keyword = "WITHDRAWAL"

# write a function to convert pdf to text using pymupdf with input as a pdf file 
# and output as a text file with the gettext() method
def pdf_to_text(input_file):
    cmd = "python3 -m fitz gettext -output tmp.txt " + input_file
    os.system(cmd)
    # print(f"Text saved to tmp.txt.")

# write a function to search for a keyword in the text file and return the float value after the keyword 
# and also return the string in the next line    
def convert_A(output_file):
    status = 200
    try:
        columns_json = read_json("columns.json")
    except Exception as e:
        status = 404
        return status, "Error in reading columns.json file."

    deposit_json = {key: [] for key in columns_json[deposit_search_keyword].keys()}
    withdrawal_json = {key: [] for key in columns_json[withdrawal_search_keyword].keys()}

    try:
        file = open("tmp.txt", 'r')
    except FileNotFoundError:
        status = 404
        return status, "PDF to text converted file not found."

    lines = file.readlines()

    for i, line in enumerate(lines):
        if deposit_search_keyword in line:
            try:
                value = line.split(deposit_search_keyword)[-1].strip()
                value = float(''.join(ch for ch in value if ch.isdecimal() or ch == '.'))

                line_without_spaces = ' '.join(lines[i+1].split())
                line_without_spaces = line_without_spaces.replace(" ", "")
                only_alphabets = "".join(char for char in line_without_spaces if char.isalpha())
                if only_alphabets == deposit_search_keyword:
                    next_line = only_alphabets
                else:
                    next_line = lines[i+1]
                    next_line = ' '.join(next_line.split())

                key = [k for k, v in enumerate(list(columns_json['DEPOSIT'].values())) if next_line in v]
                if len(key) > 0:
                    deposit_json[list(columns_json['DEPOSIT'].keys())[key[0]]].append(value)
                else:
                    deposit_json["OTHER AMOUNTS"].append(value)
                    deposit_json["OTHER VENDORS"].append(next_line)
            except Exception as e:
                status = 400
                continue

        elif withdrawal_search_keyword in line:
            try:
                value = line.split(withdrawal_search_keyword)[-1].strip()
                value = float(''.join(ch for ch in value if ch.isdecimal() or ch == '.'))

                line_without_spaces = ' '.join(lines[i+1].split())
                line_without_spaces = line_without_spaces.replace(" ", "")
                only_alphabets = "".join(char for char in line_without_spaces if char.isalpha())
                if only_alphabets == withdrawal_search_keyword:
                    next_line = only_alphabets
                else:
                    next_line = lines[i+1]
                    next_line = ' '.join(next_line.split())

                key = [k for k, v in enumerate(list(columns_json['WITHDRAWAL'].values())) if next_line in v]
                if len(key) > 0:
                    withdrawal_json[list(columns_json['WITHDRAWAL'].keys())[key[0]]].append(value)
                else:
                    withdrawal_json["OTHER AMOUNTS"].append(value)
                    withdrawal_json["OTHER VENDORS"].append(next_line)
            except Exception as e:
                status = 400
                continue

    file.close()
    os.system("rm tmp.txt")
    convert_status = json_to_excel(deposit_json, withdrawal_json, output_file)
    if convert_status == 200:
        msg = "Conversion successful."
        return status, msg
    else:
        msg = "Conversion failed."
        return convert_status, msg

def json_to_excel(deposit_json, withdrawal_json, filename):
    try:
        df_deposit = pd.DataFrame.from_dict(deposit_json, orient='index').transpose()
        df_withdrawal = pd.DataFrame.from_dict(withdrawal_json, orient='index').transpose()
        df_deposit.loc['TOTAL']= df_deposit.sum(skipna=True)
        df_deposit.loc['TOTAL', df_deposit.columns[-1]] = None
        df_withdrawal.loc['TOTAL']= df_withdrawal.sum(skipna=True)
        df_withdrawal.loc['TOTAL', df_withdrawal.columns[-1]] = None

        deposit_sum = [sum(deposit_json[a]) for a in deposit_json.keys() if a != "OTHER VENDORS"]
        deposit_sum = sum(deposit_sum)
        df_deposit.loc['ROOM_REVENUE_TOTAL', df_deposit.columns[0]] = deposit_sum

        withdrawal_sum = [sum(withdrawal_json[a]) for a in withdrawal_json.keys() if a != "OTHER VENDORS"]
        withdrawal_sum = sum(withdrawal_sum)
        df_withdrawal.loc['WITHDRAWAL_TOTAL', df_withdrawal.columns[0]] = withdrawal_sum

        writer = pd.ExcelWriter(filename, engine='xlsxwriter')   
        df_deposit.to_excel(writer, startrow=0, startcol=0)   
        df_withdrawal.to_excel(writer, startrow=len(df_deposit)+3, startcol=0) 
        writer.close()

        return 200
    except Exception as e:
        return 500, e

# write a function to read a json file and create 2 dataframes from it
def read_json(filename):
    json_data = None
    try:
        with open(filename, 'r') as file:
            json_data = json.load(file)
    except FileNotFoundError:
        print(f"File {filename} not found.")
        exit(1)
    return json_data

# Define a function to check if a file has an allowed extension
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def home():
    os.system('rm -rf *.xlsx')
    os.system('rm -rf uploads/*.pdf')
    return render_template('index.html')

# Define the route for the upload page
@app.route('/upload', methods=['POST'])
def upload_file():
    if request.method == 'POST':
        # Check if a file was uploaded
        print(request.files)
        if 'file' not in request.files:
            # return redirect(request.url)
            return "No file uploaded."
        file = request.files['file']
        # Check if the file has an allowed extension
        if file and allowed_file(file.filename):
            # Secure the filename to prevent any malicious activity
            filename = secure_filename(file.filename)
            # Save the file to the upload directory
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            to_convert_filename = UPLOAD_FOLDER + '/' + filename
            pdf_to_text(to_convert_filename)
            output_file = to_convert_filename.split("/")[-1].split(".pdf")[0] + ".xlsx"
            ret, msg = convert_A(output_file)
            return render_template('download.html', filename=output_file)
        else:
            return "File not allowed."
    return render_template('index.html')

@app.route('/download/<filename>')
def download(filename):
    print("in post download_file")
    return send_file(filename, as_attachment=True)
#    return render_template('upload.html')

if __name__ == '__main__':
    app.run(debug=True)

# if __name__ == "__main__":
#     # construct the argument parser and parse the arguments
#     ap = argparse.ArgumentParser()
#     ap.add_argument("-i", "--input", required=True, help="Path to the input PDF file")
#     args = vars(ap.parse_args())

#     # convert pdf to text
#     pdf_to_text(args["input"])
#     output_file = args["input"].split(".pdf")[0] + ".xlsx"
#     status, msg = convert_A(output_file)