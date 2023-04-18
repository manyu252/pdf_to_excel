import os
import re
import pandas as pd
import json
from flask import Flask, render_template, request, redirect, url_for, send_file
from werkzeug.utils import secure_filename

if os.name == 'nt':
    slash = '\\'  # Windows
    python = 'python'
else:
    slash = '/'  # Linux
    python = 'python3'

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024
ALLOWED_EXTENSIONS = {'pdf'}
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

deposit_search_keyword = "DEPOSIT"
withdrawal_search_keyword = "WITHDRAWAL"

# write a function to convert pdf to text using pymupdf with input as a pdf file 
# and output as a text file with the gettext() method
def pdf_to_text(input_file):
    cmd = python + " -m fitz gettext -output tmp.txt \"" + input_file + "\""
    os.system(cmd)
    # print(f"Text saved to tmp.txt.")

def get_only_alphabets(line):
    line_without_spaces = ' '.join(line.split())
    line_without_spaces = line_without_spaces.replace(" ", "")
    only_alphabets = "".join(char for char in line_without_spaces if char.isalpha())
    return only_alphabets

def remove_last_digits(line):
    updated_line = line
    space = 0
    for ind in range(len(line)-1, 0, -1):
        if space == 4:
            break
        if line[ind].isnumeric() or line[ind] == '.':
            updated_line = line[:ind]
        else:
            space = space + 1
    return updated_line

def remove_date(line):
    updated_line = line
    for ind in range(len(line)):
        if line[ind].isalpha():
            break
        if line[ind].isnumeric() or line[ind].isspace() or line[ind] == '/':
            updated_line = updated_line[1:]
    return updated_line

def check_withdrawal(line):
    ind = len(line) - 2
    if line[ind] == '-':
        if line[ind-1].isnumeric():
            if line[ind-3] == '.':
                return True, 0
    elif line[ind].isnumeric():
                return False, 0
    return False, -1

def get_account_name(line):
    updated_line = remove_date(line)
    updated_line = updated_line.split('     ')[0]
    return updated_line

# write a function to get the float value at the end of the string
def get_float_value(line):
    try:
        updated_line = remove_date(line)
        updated_line = updated_line.split('     ')[-1]
        updated_line = updated_line.replace(" ", "")
        updated_line = updated_line.replace(",", "")
        updated_line = updated_line.replace("(", "")
        updated_line = updated_line.replace(")", "")
        updated_line = updated_line.replace("-", "")
        return float(updated_line)
    except Exception as e:
        print("Error in converting to float: ", e)
        return None


def convert_CIB(output_file):
    status = 200
    try:
        columns_json = read_json("columns.json")
    except Exception as e:
        status = 404
        return status, "Error in reading columns.json file."

    deposit_json = {key: [] for key in columns_json[deposit_search_keyword].keys()}
    withdrawal_json = {key: [] for key in columns_json[withdrawal_search_keyword].keys()}
    loan_list = columns_json[deposit_search_keyword]["LOAN"]
    loan_lines = []

    try:
        file = open("tmp.txt", 'r')
    except FileNotFoundError:
        status = 404
        return status, "PDF to text converted file not found."

    lines = file.readlines()

    # This is for adding the loan amount to the deposit json
    for i, line in enumerate(lines):
        stored_line = line
        updated_line = remove_last_digits(stored_line)
        updated_line = remove_date(updated_line)

        account_name = ' '.join(updated_line.split())
        if account_name == '':
            continue

        for val in loan_list:
            if val in account_name:
                try:
                    value = lines[i].split(account_name)[-1].strip()
                    value = float(''.join(ch for ch in value if ch.isdecimal() or ch == '.'))
                    deposit_json["LOAN"].append(value)
                    loan_lines.append(lines[i])
                except Exception as e:
                    print("Error in converting loan format: ", e)
                    status = 400
                    continue

    is_deposit = False
    is_withdrawal = False
    date_check = re.compile(r'\b\d{1,2}\s*/\s*\d{1,2}\b')

    # This is for adding the deposit and withdrawal amount to the json
    for i, line in enumerate(lines):
        if "CHECKS IN NUMBER ORDER" in line:
            break

        if "Page" in line:
            is_deposit = False
            is_withdrawal = False
            continue

        if "DEPOSITS AND ADDITIONS" in line:
            is_deposit = True
            is_withdrawal = False
            continue

        if "CHECKS AND WITHDRAWALS" in line:
            is_deposit = False
            is_withdrawal = True
            continue

        if not is_deposit and not is_withdrawal:
            continue

        date_match = date_check.search(line)
        if not date_match:
            continue

        if lines[i] in loan_lines:
            loan_lines.remove(lines[i])
            continue

        account = remove_last_digits(lines[i])
        account = remove_date(account)
        account_name = ' '.join(account.split())

        if is_deposit:
            try:
                value = line.split(account)[-1].strip()
                value = float(''.join(ch for ch in value if ch.isdecimal() or ch == '.'))

                key = None
                for k, v in enumerate(list(columns_json[deposit_search_keyword].values())):
                    if key is not None:
                        break
                    for i in v:
                        if i in account_name:
                            key = k
                            break
                if key is not None:
                    deposit_json[list(columns_json[deposit_search_keyword].keys())[key]].append(value)
                else:
                    deposit_json["OTHER AMOUNTS"].append(value)
                    deposit_json["OTHER VENDORS"].append(account_name)
            except Exception as e:
                status = 400
                print("Error in converting deposit format: ", e)
                continue

        elif is_withdrawal:
            try:
                value = line.split(account)[-1].strip()
                value = float(''.join(ch for ch in value if ch.isdecimal() or ch == '.'))

                key = None
                for k, v in enumerate(list(columns_json[withdrawal_search_keyword].values())):
                    if key is not None:
                        break
                    for i in v:
                        if i in account_name:
                            key = k
                            break
                if key is not None:
                    withdrawal_json[list(columns_json[withdrawal_search_keyword].keys())[key]].append(value)
                else:
                    withdrawal_json["OTHER AMOUNTS"].append(value)
                    withdrawal_json["OTHER VENDORS"].append(account_name)
            except Exception as e:
                status = 400
                print("Error in converting withdrawal format: ", e)
                continue

    file.close()
    convert_status = json_to_excel(deposit_json, withdrawal_json, output_file)
    if convert_status == 200:
        msg = "Conversion successful."
        return status, msg
    else:
        msg = "Conversion failed."
        return convert_status, msg

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
    loan_list = columns_json[deposit_search_keyword]["LOAN"]
    loan_lines = []

    try:
        file = open("tmp.txt", 'r')
    except FileNotFoundError:
        status = 404
        return status, "PDF to text converted file not found."

    lines = file.readlines()

    # This is for adding the loan amount to the deposit json
    for i, line in enumerate(lines):
        stored_line = line
        updated_line = remove_last_digits(stored_line)
        updated_line = remove_date(updated_line)

        account_name = ' '.join(updated_line.split())
        if account_name == '':
            continue

        for val in loan_list:
            if val in account_name:
                try:
                    value = lines[i].split(account_name)[-1].strip()
                    value = float(''.join(ch for ch in value if ch.isdecimal() or ch == '.'))
                    deposit_json["LOAN"].append(value)
                    loan_lines.append(lines[i])
                except Exception as e:
                    print("Error in converting loan format: ", e)
                    status = 400
                    continue

    # This is for adding the deposit and withdrawal amount to the json
    for i, line in enumerate(lines):
        check_paymentech = ' '.join(line.split())
        if check_paymentech.lower() == "account title: sunrise hospitality, llc":
            columns_json[deposit_search_keyword]["VISA/MC"].remove("DEPOSIT PAYMENTECH")

        if deposit_search_keyword in line:
            try:
                value = line.split(deposit_search_keyword)[-1].strip()
                value = float(''.join(ch for ch in value if ch.isdecimal() or ch == '.'))

                only_alphabets = get_only_alphabets(lines[i+1])
                if only_alphabets == deposit_search_keyword:
                    next_line = only_alphabets
                elif lines[i+1] in loan_lines:
                    next_line = only_alphabets
                    loan_lines.remove(lines[i+1])
                else:
                    next_line = lines[i+1]
                    next_line = ' '.join(next_line.split())

                key = None
                for k, v in enumerate(list(columns_json[deposit_search_keyword].values())):
                    if key is not None:
                        break
                    for i in v:
                        if i in next_line:
                            key = k
                            break
                if key is not None:
                    deposit_json[list(columns_json[deposit_search_keyword].keys())[key]].append(value)
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

                only_alphabets = get_only_alphabets(lines[i+1])
                if only_alphabets == withdrawal_search_keyword:
                    next_line = only_alphabets
                else:
                    next_line = lines[i+1]
                    next_line = ' '.join(next_line.split())

                key = None
                for k, v in enumerate(list(columns_json[withdrawal_search_keyword].values())):
                    if key is not None:
                        break
                    for i in v:
                        if i in next_line:
                            key = k
                            break
                if key is not None:
                    withdrawal_json[list(columns_json[withdrawal_search_keyword].keys())[key]].append(value)
                else:
                    withdrawal_json["OTHER AMOUNTS"].append(value)
                    withdrawal_json["OTHER VENDORS"].append(next_line)
            except Exception as e:
                status = 400
                continue

    file.close()
    convert_status = json_to_excel(deposit_json, withdrawal_json, output_file)
    if convert_status == 200:
        msg = "Conversion successful."
        return status, msg
    else:
        msg = "Conversion failed."
        return convert_status, msg

def convert_DIP(output_file):
    status = 200
    try:
        columns_json = read_json("columns.json")
    except Exception as e:
        status = 404
        return status, "Error in reading columns.json file."

    deposit_json = {key: [] for key in columns_json[deposit_search_keyword].keys()}
    withdrawal_json = {key: [] for key in columns_json[withdrawal_search_keyword].keys()}
    loan_list = columns_json[deposit_search_keyword]["LOAN"]

    try:
        file = open("tmp.txt", 'r')
    except FileNotFoundError:
        status = 404
        return status, "PDF to text converted file not found."

    transaction_pattern = re.compile(r'\d+\.\d+') # This is for getting the transaction amount
    lines = file.readlines()

    # This is for adding the loan amount to the deposit json
    for i, line in enumerate(lines):
        line = line.replace(',', '')
        match = transaction_pattern.search(line)
        if not match:
            continue

        transaction_amount = match.group()
        transaction_amount = float(transaction_amount)

        account_name = get_account_name(line)

        # next_line = ' '.join(line.split())
        was_loan = False

        for val in loan_list:
            if val in account_name:
                try:
                    deposit_json["LOAN"].append(transaction_amount)
                    was_loan = True
                except Exception as e:
                    print("Error in converting loan format: ", e)
                    status = 400
                    continue

        if was_loan:
            continue

        if account_name == deposit_search_keyword:
            match = transaction_pattern.search(lines[i+1])
            if not match and lines[i+1] != '':
                account_name = ' '.join(lines[i+1].split())

            key = None
            for k, v in enumerate(list(columns_json[deposit_search_keyword].values())):
                if key is not None:
                    break
                for i in v:
                    if i in account_name:
                        key = k
                        break
            if key is not None:
                deposit_json[list(columns_json[deposit_search_keyword].keys())[key]].append(transaction_amount)
            else:
                deposit_json["OTHER AMOUNTS"].append(transaction_amount)
                deposit_json["OTHER VENDORS"].append(account_name)

        elif account_name == withdrawal_search_keyword:
            match = transaction_pattern.search(lines[i+1])
            if not match and lines[i+1] != '':
                account_name = ' '.join(lines[i+1].split())

            key = None
            for k, v in enumerate(list(columns_json[withdrawal_search_keyword].values())):
                if key is not None:
                    break
                for i in v:
                    if i in account_name:
                        key = k
                        break
            if key is not None:
                withdrawal_json[list(columns_json[withdrawal_search_keyword].keys())[key]].append(transaction_amount)
            else:
                withdrawal_json["OTHER AMOUNTS"].append(transaction_amount)
                withdrawal_json["OTHER VENDORS"].append(account_name)

    file.close()
    convert_status = json_to_excel(deposit_json, withdrawal_json, output_file)
    if convert_status == 200:
        msg = "Conversion successful."
        return status, msg
    else:
        msg = "Conversion failed."
        return convert_status, msg

def convert_TPS(output_file):
    status = 200
    try:
        columns_json = read_json("columns.json")
    except Exception as e:
        status = 404
        return status, "Error in reading columns.json file."

    deposit_json = {key: [] for key in columns_json[deposit_search_keyword].keys()}
    withdrawal_json = {key: [] for key in columns_json[withdrawal_search_keyword].keys()}
    loan_list = columns_json[deposit_search_keyword]["LOAN"]

    try:
        file = open("tmp.txt", 'r')
    except FileNotFoundError:
        status = 404
        return status, "PDF to text converted file not found."

    date_check = re.compile(r'\b\d{2}/\d{2}\b')

    lines = file.readlines()
    is_deposit = False
    is_withdrawal = False
    counts = 0

    # This is for adding the loan amount to the deposit json
    for i, line in enumerate(lines):
        if "Deposits, credits and interest" in line:
            counts += 1
            if counts > 2:
                is_deposit = True
                is_withdrawal = False
        elif "Other withdrawals, debits and service charges" in line:
            counts += 1
            if counts > 2:
                is_withdrawal = True
                is_deposit = False

        if not is_deposit and not is_withdrawal:
            continue

        date_match = date_check.search(line)
        if not date_match:
            continue

        line = line.replace(',', '')
        transaction_amount = get_float_value(line)
        if transaction_amount is None:
            continue

        account = remove_last_digits(lines[i])
        account = remove_date(account)
        account_name = ' '.join(account.split())

        if is_deposit:
            if "Total deposits, credits and interest" in line:
                is_deposit = False
                continue

            key = None
            for k, v in enumerate(list(columns_json[deposit_search_keyword].values())):
                if key is not None:
                    break
                for i in v:
                    if i in account_name:
                        key = k
                        break
            if key is not None:
                deposit_json[list(columns_json[deposit_search_keyword].keys())[key]].append(transaction_amount)
            else:
                deposit_json["OTHER AMOUNTS"].append(transaction_amount)
                deposit_json["OTHER VENDORS"].append(account_name)

        elif is_withdrawal:
            if "Total other withdrawals, debits and service charges" in line:
                is_withdrawal = False
                continue

            key = None
            for k, v in enumerate(list(columns_json[withdrawal_search_keyword].values())):
                if key is not None:
                    break
                for i in v:
                    if i in account_name:
                        key = k
                        break
            if key is not None:
                withdrawal_json[list(columns_json[withdrawal_search_keyword].keys())[key]].append(transaction_amount)
            else:
                withdrawal_json["OTHER AMOUNTS"].append(transaction_amount)
                withdrawal_json["OTHER VENDORS"].append(account_name)

    file.close()
    convert_status = json_to_excel(deposit_json, withdrawal_json, output_file)
    if convert_status == 200:
        msg = "Conversion successful."
        return status, msg
    else:
        msg = "Conversion failed."
        return convert_status, msg

# Function to convert HIP and SIB pdf types
def convert_HIP(output_file):
    status = 200
    try:
        columns_json = read_json("columns.json")
    except Exception as e:
        status = 404
        return status, "Error in reading columns.json file."

    deposit_json = {key: [] for key in columns_json[deposit_search_keyword].keys()}
    withdrawal_json = {key: [] for key in columns_json[withdrawal_search_keyword].keys()}
    loan_list = columns_json[deposit_search_keyword]["LOAN"]

    try:
        file = open("tmp.txt", 'r')
    except FileNotFoundError:
        status = 404
        return status, "PDF to text converted file not found."

    transaction_pattern = re.compile(r'\d+\.\d+') # This is for getting the transaction amount
    lines = file.readlines()
    start = False

    # This is for adding the loan amount to the deposit json
    for i, line in enumerate(lines):
        if "Balance Last Statement" in line:
            start = True
            continue
        elif "Balance This Statement" in line:
            start = False

        if not start:
            continue

        if "Check #" in line:
            continue

        line = line.replace(',', '')
        match = transaction_pattern.search(line)
        if not match:
            continue

        transaction_amount = match.group()
        transaction_start_loc = match.start()
        transaction_end_loc = match.end()
        transaction_amount = float(transaction_amount)

        account_name = get_account_name(line)

        is_deposit = False
        is_withdrawal = False
        fraction = transaction_end_loc / len(line)

        if fraction >= 0.85:
            is_deposit = True
            is_withdrawal = False
        elif fraction <= 0.85:
            is_deposit = False
            is_withdrawal = True

        if is_deposit:
            key = None
            for k, v in enumerate(list(columns_json[deposit_search_keyword].values())):
                if key is not None:
                    break
                for i in v:
                    if i in account_name:
                        key = k
                        break
            if key is not None:
                deposit_json[list(columns_json[deposit_search_keyword].keys())[key]].append(transaction_amount)
            else:
                deposit_json["OTHER AMOUNTS"].append(transaction_amount)
                deposit_json["OTHER VENDORS"].append(account_name)

        elif is_withdrawal:
            key = None
            for k, v in enumerate(list(columns_json[withdrawal_search_keyword].values())):
                if key is not None:
                    break
                for i in v:
                    if i in account_name:
                        key = k
                        break
            if key is not None:
                withdrawal_json[list(columns_json[withdrawal_search_keyword].keys())[key]].append(transaction_amount)
            else:
                withdrawal_json["OTHER AMOUNTS"].append(transaction_amount)
                withdrawal_json["OTHER VENDORS"].append(account_name)

    file.close()
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

        deposit_cols = list(deposit_json)
        deposit_cols.remove("OTHER VENDORS")
        withdrawal_cols = list(withdrawal_json)
        withdrawal_cols.remove("OTHER VENDORS")

        df_deposit.loc['TOTAL']= df_deposit[deposit_cols].sum(skipna=True)
        df_withdrawal.loc['TOTAL']= df_withdrawal[withdrawal_cols].sum(skipna=True)

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
        print("Error in converting json to excel: ", e)
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

def rem():
    try:
        os.remove("tmp.txt")
    except:
        print("Error while deleting file : tmp.txt")
        pass

    dir = os.getcwd()
    test = os.listdir(dir)
    try:
        for item in test:
            if item.endswith(".xlsx"):
                path = os.path.join(dir, item)
                os.remove(path)
    except:
        print("Error while deleting file : ", item)
        pass

    try:
        dir = UPLOAD_FOLDER
        test = os.listdir(dir)
        for item in test:
            if item.endswith(".pdf"):
                path = os.path.join(dir, item)
                os.remove(path)
    except:
        print("Error while deleting file : ", item)
        pass

# Define a function to check if a file has an allowed extension
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def home():
    rem()
    return render_template('index.html')

# Define the route for the upload page
@app.route('/upload', methods=['POST'])
def upload_file():
    if request.method == 'POST':
        # Check if a file was uploaded
        if 'file' not in request.files:
            # return redirect(request.url)
            return "No file uploaded."
        file = request.files['file']
        selected_option = request.form['option']
        if selected_option == "Choose":
            return "Choose a valid PDF type"
        # Check if the file has an allowed extension
        if file and allowed_file(file.filename):
            # Secure the filename to prevent any malicious activity
            filename = secure_filename(file.filename)
            # Save the file to the upload directory
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            to_convert_filename = os.path.join(UPLOAD_FOLDER, filename)
            pdf_to_text(to_convert_filename)
            output_file = to_convert_filename.split(slash)[-1].split(".pdf")[0] + ".xlsx"
            if selected_option == "HIP":
                status, msg = convert_HIP(output_file)
            elif selected_option == "DIP":
                status, msg = convert_DIP(output_file)
            elif selected_option == "CIB":
                status, msg = convert_CIB(output_file)
            elif selected_option == "TPS":
                status, msg = convert_TPS(output_file)
            return render_template('download.html', filename=output_file)
        else:
            return "File not allowed."
    return render_template('index.html')

@app.route('/download/<filename>')
def download(filename):
    return send_file(filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
