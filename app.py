from flask import Flask, render_template, request
import openpyxl
from openpyxl import Workbook

app = Flask(__name__)

# Route to serve the form
@app.route('/')
def form():
    return render_template('form.html')

# Route to handle form submission
@app.route('/submit', methods=['POST'])
def submit():
    # Get form data
    name = request.form['name']
    email = request.form['email']
    age = request.form['age']
    gender = request.form['gender']
    course = request.form['course']
    
    # Save data to Excel
    save_to_excel([name, email, age, gender, course])
    
    # Redirect to success page after saving
    return render_template('success.html')

# Function to save form data to an Excel file
def save_to_excel(data, filename='students_data.xlsx'):
    try:
        # Try to open the Excel file if it exists
        workbook = openpyxl.load_workbook(filename)
        sheet = workbook.active
    except FileNotFoundError:
        # Create a new Excel file if it doesn't exist
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Name", "Email", "Age", "Gender", "Course"])  # Add headers
    
    # Append the new form data
    sheet.append(data)
    
    # Save the workbook
    workbook.save(filename)
    print(f"Data saved to {filename}")

# Run the web server
if __name__ == '__main__':
    app.run(debug=True)
