from flask import Flask, render_template_string, request, redirect, url_for
import openpyxl
from openpyxl import Workbook

app = Flask(__name__)

# HTML template for the form
form_html = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Student Information Form</title>
    <style>
        body {
            font-family: 'Arial', sans-serif;
            background-color: #f4f4f4;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
        }

        .form-container {
            background-color: #fff;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 4px 10px rgba(0, 0, 0, 0.1);
            width: 400px;
        }

        .form-container h2 {
            text-align: center;
            margin-bottom: 20px;
            color: #333;
        }

        label {
            font-weight: bold;
            margin-top: 10px;
            display: block;
        }

        input[type="text"], input[type="email"], input[type="number"], select {
            width: 100%;
            padding: 10px;
            margin-top: 5px;
            border: 1px solid #ccc;
            border-radius: 5px;
            box-sizing: border-box;
        }

        input[type="submit"] {
            width: 100%;
            padding: 10px;
            background-color: #28a745;
            color: #fff;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            margin-top: 20px;
            font-size: 16px;
        }

        input[type="submit"]:hover {
            background-color: #218838;
        }
    </style>
</head>
<body>
    <div class="form-container">
        <h2>Student Information</h2>
        <form action="/submit" method="POST">
            <label for="name">Full Name:</label>
            <input type="text" id="name" name="name" required>

            <label for="email">Email:</label>
            <input type="email" id="email" name="email" required>

            <label for="age">Age:</label>
            <input type="number" id="age" name="age" required>

            <label for="gender">Gender:</label>
            <select id="gender" name="gender" required>
                <option value="">Select Gender</option>
                <option value="Male">Male</option>
                <option value="Female">Female</option>
                <option value="Other">Other</option>
            </select>

            <label for="course">Course:</label>
            <input type="text" id="course" name="course" required>

            <input type="submit" value="Submit">
        </form>
    </div>
</body>
</html>
"""

# HTML for the success page
success_html = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Success</title>
    <style>
        body {
            font-family: 'Arial', sans-serif;
            background-color: #f4f4f4;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
        }

        .confirmation-container {
            background-color: #fff;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 4px 10px rgba(0, 0, 0, 0.1);
            width: 400px;
            text-align: center;
        }

        .confirmation-container h2 {
            color: #28a745;
        }

        button {
            padding: 10px 20px;
            background-color: #28a745;
            color: white;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            font-size: 16px;
        }

        button:hover {
            background-color: #218838;
        }
    </style>
</head>
<body>
    <div class="confirmation-container">
        <h2>Form Submitted Successfully!</h2>
        <p>Your data has been saved.</p>
        <button onclick="window.location.href='/'">Back to Form</button>
    </div>
</body>
</html>
"""

# Route to serve the form
@app.route('/')
def form():
    return render_template_string(form_html)

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
    return render_template_string(success_html)

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
