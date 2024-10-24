from flask import Flask, request, render_template_string
import pandas as pd
import os

app = Flask(__name__)

# Path to the Excel file
EXCEL_FILE = 'survey_results.xlsx'

# Initialize the Excel file if it doesn't exist
if not os.path.exists(EXCEL_FILE):
    # Create a DataFrame to hold the results with columns: Question, Answer, and Count
    df = pd.DataFrame(columns=["Question", "Answer", "Count"])
    # Save it as an Excel file using xlsxwriter
    df.to_excel(EXCEL_FILE, index=False, engine='xlsxwriter')

# Route to display the survey form (Home Page)
@app.route('/')
def survey_form():
    html_form = '''
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Survey Form</title>
    </head>
    <body>
        <h1>Survey Form</h1>
        <form action="/submit-survey" method="POST">
            <!-- Question 1 -->
            <label for="q1">1. How satisfied are you with our website?</label><br>
            <input type="radio" id="q1_very_satisfied" name="q1" value="Very Satisfied">
            <label for="q1_very_satisfied">Very Satisfied</label><br>
            <input type="radio" id="q1_satisfied" name="q1" value="Satisfied">
            <label for="q1_satisfied">Satisfied</label><br>
            <input type="radio" id="q1_neutral" name="q1" value="Neutral">
            <label for="q1_neutral">Neutral</label><br>
            <input type="radio" id="q1_dissatisfied" name="q1" value="Dissatisfied">
            <label for="q1_dissatisfied">Dissatisfied</label><br>
            <input type="radio" id="q1_very_dissatisfied" name="q1" value="Very Dissatisfied">
            <label for="q1_very_dissatisfied">Very Dissatisfied</label><br><br>

            <!-- Question 2 -->
            <label for="q2">2. How often do you visit our website?</label><br>
            <input type="radio" id="q2_daily" name="q2" value="Daily">
            <label for="q2_daily">Daily</label><br>
            <input type="radio" id="q2_weekly" name="q2" value="Weekly">
            <label for="q2_weekly">Weekly</label><br>
            <input type="radio" id="q2_monthly" name="q2" value="Monthly">
            <label for="q2_monthly">Monthly</label><br>
            <input type="radio" id="q2_rarely" name="q2" value="Rarely">
            <label for="q2_rarely">Rarely</label><br><br>

            <!-- Question 3 -->
            <label for="q3">3. What would you like to see improved?</label><br>
            <input type="checkbox" id="q3_design" name="q3" value="Design">
            <label for="q3_design">Design</label><br>
            <input type="checkbox" id="q3_content" name="q3" value="Content">
            <label for="q3_content">Content</label><br>
            <input type="checkbox" id="q3_performance" name="q3" value="Performance">
            <label for="q3_performance">Performance</label><br>
            <input type="checkbox" id="q3_usability" name="q3" value="Usability">
            <label for="q3_usability">Usability</label><br><br>

            <!-- Submit button -->
            <input type="submit" value="Submit">
        </form>
    </body>
    </html>
    '''
    return render_template_string(html_form)

# Route to handle form submission
@app.route('/submit-survey', methods=['POST'])
def submit_survey():
    # Get the form data
    q1 = request.form.get('q1')
    q2 = request.form.get('q2')
    q3 = request.form.getlist('q3')  # Checkboxes return a list

    # Create a list of responses
    responses = [
        {"Question": "How satisfied are you with our website?", "Answer": q1},
        {"Question": "How often do you visit our website?", "Answer": q2}
    ]

    # Add responses for each checkbox entry in q3
    for answer in q3:
        responses.append({"Question": "What would you like to see improved?", "Answer": answer})

    # Load the existing Excel file using openpyxl (for reading)
    df = pd.read_excel(EXCEL_FILE, engine='openpyxl')

    # Ensure columns exist in the file
    if 'Question' not in df.columns or 'Answer' not in df.columns or 'Count' not in df.columns:
        return "Error: Excel file does not have the required columns."

    # Update the tally for each response
    for response in responses:
        # Check if the response already exists in the file
        existing_entry = df[(df['Question'] == response['Question']) & (df['Answer'] == response['Answer'])]

        if not existing_entry.empty:
            # If it exists, increment the count
            df.loc[existing_entry.index, 'Count'] += 1
        else:
            # If not, add a new row with a count of 1
            new_entry = pd.DataFrame([{"Question": response['Question'], "Answer": response['Answer'], "Count": 1}])
            df = pd.concat([df, new_entry], ignore_index=True)

    # Write the updated DataFrame back to the Excel file using xlsxwriter (for writing)
    df.to_excel(EXCEL_FILE, index=False, engine='xlsxwriter')

    return "Survey submitted successfully!"

# Start the Flask app
if __name__ == '__main__':
    app.run(debug=True)
