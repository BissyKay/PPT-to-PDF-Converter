from flask import Flask, request, send_file, render_template, flash, redirect, url_for
from werkzeug.utils import secure_filename
from pptx import Presentation
import os

# Initialize Flask app
app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = "uploads"
app.config["CONVERTED_FOLDER"] = "converted"
app.config["ALLOWED_EXTENSIONS"] = {"pptx"}
app.secret_key = "your_secret_key_here"  # For flash messages

# Ensure folders exist
os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
os.makedirs(app.config["CONVERTED_FOLDER"], exist_ok=True)

# Function to check allowed file types
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in app.config["ALLOWED_EXTENSIONS"]

# Home route (Upload Page)
@app.route("/")
def home():
    return render_template("index.html")

# File Upload & Conversion Route
@app.route("/convert", methods=["POST"])
def convert():
    if "file" not in request.files:
        flash("No file uploaded!", "error")
        return redirect(url_for("home"))

    file = request.files["file"]
    if file.filename == "":
        flash("No file selected!", "error")
        return redirect(url_for("home"))

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        file.save(filepath)

        pdf_filename = filename.rsplit(".", 1)[0] + ".pdf"
        pdf_filepath = os.path.join(app.config["CONVERTED_FOLDER"], pdf_filename)

        try:
            presentation = Presentation(filepath)
            presentation.save(pdf_filepath)
            flash("Conversion successful! Your PDF is ready.", "success")
            return send_file(pdf_filepath, as_attachment=True)
        except Exception as e:
            flash(f"Conversion failed: {str(e)}", "error")
            return redirect(url_for("home"))

    else:
        flash("Invalid file type! Only .pptx files are allowed.", "error")
        return redirect(url_for("home"))

if __name__ == "__main__":
    app.run(debug=True)
