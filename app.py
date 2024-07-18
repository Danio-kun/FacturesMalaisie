from flask import Flask, request, render_template, redirect, url_for
import os
import pandas as pd
from datetime import datetime

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'

# Ensure the upload directory exists
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

# Ensure the data directory exists
data_folder = 'data/'
if not os.path.exists(data_folder):
    os.makedirs(data_folder)

# Path to the Excel file
excel_path = os.path.join(data_folder, 'Récaptulatif dépenses Mobilité Internationale Daniel FINDRAMA.xlsx')

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    nature_depense = request.form['nature_depense']
    date_depense = request.form['date_depense']
    somme_depense = request.form['somme_depense']
    file = request.files['file']

    # If date is not provided, use today's date
    if not date_depense:
        date_depense = datetime.today()
    else:
        date_depense = datetime.strptime(date_depense, '%Y-%m-%d')

    # Determine the month and create the directory if it does not exist
    month_folder = os.path.join(app.config['UPLOAD_FOLDER'], date_depense.strftime('%B').lower())
    if not os.path.exists(month_folder):
        os.makedirs(month_folder)

    # Save the file
    file_path = os.path.join(month_folder, file.filename)
    file.save(file_path)

    # Update the Excel file
    if os.path.exists(excel_path):
        df = pd.read_excel(excel_path)
    else:
        df = pd.DataFrame(columns=['Nature Dépense', 'Date', 'Chemin', 'Fichier facture', 'Somme Dépense'])

    new_row = pd.DataFrame([{
        'Nature Dépense': nature_depense,
        'Date': date_depense.strftime('%Y-%m-%d'),
        'Chemin': file_path,
        'Fichier facture': file.filename,
        'Somme Dépense': somme_depense
    }])

    df = pd.concat([df, new_row], ignore_index=True)
    df.to_excel(excel_path, index=False)

    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
