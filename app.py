from flask import Flask, render_template, request, send_file
from flask_bootstrap import Bootstrap5
from docx import Document
from python_docx_replace import docx_replace
import os
from datetime import datetime

app = Flask(__name__)
app.config['MODULI_FOLDER'] = 'module'
app.config['OUTPUT_FOLDER'] = 'output'
bootstrap = Bootstrap5(app)

# Assicurati che la cartella di output esista
if not os.path.exists(app.config['OUTPUT_FOLDER']):
    os.makedirs(app.config['OUTPUT_FOLDER'])

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/fviaggio', methods=['GET','POST'])

def fviaggio():

    form = request.form

    if request.method == 'POST':
        # Ottieni i dati dal modulo
        
        mydict = {
            'qual': request.form['qual'],
            'name': request.form['name'],
            'surname': request.form['surname'],
            'perid': request.form['perid'],
            'nr_fvg': request.form['nr_fvg'],
            'motive': request.form['motive'],
            'cod': request.form['mission_code'],
            'cod_off': request.form['office_code'],
            'dest': request.form['destination'],
            'data': datetime.strptime(request.form['date'],'%Y-%m-%d').strftime('%d-%m-%Y'),
            'month': datetime.strptime(request.form['date'],'%Y-%m-%d').strftime('%m'),
            'year': datetime.strptime(request.form['date'],'%Y-%m-%d').strftime('%Y'),
            'h_serv': request.form['h_serv'],
            'data_fvg': datetime.strptime(request.form['data_fvg'],'%Y-%m-%d').strftime('%d-%m-%Y'),
            'from1': request.form['from1'],
            'to1': request.form['to1'],
            'from2': request.form['from2'],
            'to2': request.form['to2'],
            'from3': request.form['from3'],
            'to3': request.form['to3'],
            'from4': request.form['from4'],
            'to4': request.form['to4'],
            'from5': request.form['from5'],
            'to5': request.form['to5'],
            'from6': request.form['from6'],
            'to6': request.form['to6'],
            'yn1': '[X]' if request.form['canteen'] == 'YES' else '[ ]',
            'yn2': '[X]' if request.form['canteen'] == 'NO' else '[ ]',
            'yn3': '[X]' if request.form['lodge'] == 'YES' else '[ ]',
            'yn4': '[X]' if request.form['lodge'] == 'NO' else '[ ]',
            'canteen': 'PRESENTE' if request.form['canteen'] == 'YES' else 'NON PRESENTE',
            'accomodation': 'PRESENTE' if request.form['lodge'] == 'YES' else 'NON PRESENTE'
            }

        # Carica il template Word
        template_path = os.path.join(app.config['MODULI_FOLDER'], 'f_viaggio.docx')
        doc = Document(template_path)

        # Sostituisci i segnaposto con i dati del modulo
        docx_replace(doc, **mydict)
        # Salva il documento modificato 
        output_filename = "{}_{}_foglio_viaggio_{}.docx".format(mydict['surname'], mydict['name'], datetime.now().strftime("%Y%m%d_%H%M%S"))
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
        doc.save(output_path)
        # Invia il file come download  
        return send_file(output_path, as_attachment=True)
    
    # Se non è una richiesta POST, mostra il modulo 
    return render_template('fviaggio.html', form=form)
       