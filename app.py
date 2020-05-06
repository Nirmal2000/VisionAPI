import os
#import magic
import urllib.request
from flask import Flask, flash, request, redirect, render_template,send_file,make_response,Response
from werkzeug.utils import secure_filename
import vision_api_demo
from tempfile import NamedTemporaryFile
import flask_excel as excel
from openpyxl.writer.excel import save_virtual_workbook
from io import BytesIO
UPLOAD_FOLDER = './upload_img/'

app = Flask(__name__)
app.secret_key = "secret key"
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024


ALLOWED_EXTENSIONS = set(['txt', 'pdf', 'png', 'jpg', 'jpeg', 'gif'])

def allowed_file(filename):
	return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
	
@app.route('/')
def upload_form():    
	result = './templates/result.jpg'
	example = './templates/9.jpg'
	return render_template('license.html')

@app.route('/download')
def download_form():
    return render_template('download.html')



@app.route('/download',methods=['POST'])
def download_excel():
    filelist = [ f for f in os.listdir(UPLOAD_FOLDER) if f.endswith(".jpg") ]
    for f in filelist:
        os.remove(os.path.join(UPLOAD_FOLDER, f))

    return send_file('license.xlsx',as_attachment=True)

@app.route('/', methods=['POST'])
def upload_file():
	if request.method == 'POST':
        # check if the post request has the files part
		if 'files[]' not in request.files:
			flash('No file part')
			return redirect(request.url)
		files = request.files.getlist('files[]')
		# files = request.files['files']
		# for file in files:
		# 	if file and allowed_file(file.filename):
		# 		print(file)
		# 		filename = secure_filename(file.filename)
		# 		file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        
		wb = vision_api_demo.start_detection(files)
		print(wb)
		# myfile = BytesIO()		
		# myfile.write(save_virtual_workbook(wb))		
		response_string = '' 
		for line in wb:
			for word in line:
				response_string+=word+','
			response_string = response_string[:-1]
			response_string+='\n'
		
		response = Response(response_string,content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={"Contentdisposition":"attachment; filename=" + "a.xlsx"})
		return response
		


if __name__ == "__main__":
    app.run()