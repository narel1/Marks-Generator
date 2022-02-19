# Name: Anushkha Singh
# Roll No.: 1901ME72

# Filename: app.py : Execute this file then open the browser and type the URL: http://127.0.0.1:5000/ to use the web application

# Import relevant libraries
import json
from werkzeug.utils import secure_filename
from flask import Flask, render_template, request, flash, session, redirect, url_for
import os
from work_main import generate_marksheet, concise_marksheet, Send_email, error_handling
import pandas as pd
from wtforms import Form, FloatField, validators

# Initialize Flask object to create a web page and route to web pages
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = "sample_input"
app.secret_key = "hello"

# Flask app routes the user to the index page of the web GUI
@app.route('/', methods=['GET', 'POST'])
def index():

	# Use POST method to fetch the information entered in the form
	if request.method == "POST":
		req = request.form

		# Fetch the value of marks as per the correct answers and negative marks
		pos = req.get("correct")
		neg = req.get("wrong")

		# Create the input folder if it doesn't exist
		if not os.path.isdir(app.config['UPLOAD_FOLDER']):
			os.mkdir(app.config['UPLOAD_FOLDER'])

		# Save the files uploaded by user to the sample_input folder 
		f1 = request.files["upload-file1"]

		x= None

        # Error handling
		if error_handling(f1, app.config['UPLOAD_FOLDER']) != None:
			x = error_handling(f1, app.config['UPLOAD_FOLDER'])

		f2 = request.files["upload-file2"]

		# Error handling 
		if error_handling(f2, app.config['UPLOAD_FOLDER']) != None:
			x = error_handling(f2, app.config['UPLOAD_FOLDER'])

        # Throw the error message
		if x != None:
			flash(x)
			return redirect(request.url)

        # Complete path for the uploaded files
		path1 = os.path.join(app.config['UPLOAD_FOLDER'], f1.filename)
		path2 = os.path.join(app.config['UPLOAD_FOLDER'], f2.filename)

		# Actions performed when user clicks on the 3 buttons of the GUI 
		if (pos and neg and f1 and f2) :

			if request.form["action"] == "Generate Roll number wise Marksheet":
				result = generate_marksheet(path1, path2, pos, neg)

			elif request.form["action"] == "Generate Concise Marksheet with Roll Num, Obtained Marks, marks after negative":
				result = concise_marksheet(path1, path2, pos, neg)

			elif request.form["action"] == "Send Email":
				result = Send_email()
			
			else:
				result = None

		else:

			result = None

			if request.form["action"] == "Generate Roll number wise Marksheet":
				flash("Please select files and enter the values!!")

			elif request.form["action"] == "Generate Concise Marksheet with Roll Num, Obtained Marks, marks after negative":
				flash("Please select files and enter the values!!")

			elif request.form["action"] == "Send Email":
				result = Send_email()

		# Flash error messages to make sure program runs smoothly
		if result:
			flash(result)
			return redirect(request.url)

		# Renders the main page of GUI
		return redirect(request.url)

	# Renders the main page of GUI
	return render_template('view.html')

if __name__ == '__main__':

	# Call the run function of flask app and enabling debugging in case of any errors
    app.run(debug=True)