import os
#import magic
import urllib.request
from flask import Flask, flash, request, redirect, render_template,send_file
from werkzeug.utils import secure_filename
import vision_api_demo

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

		for file in files:
			if file and allowed_file(file.filename):
				filename = secure_filename(file.filename)
				file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        
		vision_api_demo.start_detection()        
		return redirect('/download')

if __name__ == "__main__":
    app.run()