from flask import Flask, request, render_template
import pandas as pd
import os

app = Flask(__name__)

# Paths to Excel files
SURVEY_FILE = 'survey_results.xlsx'
INFO_FILE = 'user_info.xlsx'

# Initialize the Excel files if they don't exist
if not os.path.exists(SURVEY_FILE):
    df = pd.DataFrame(columns=["Question", "Answer"])
    df.to_excel(SURVEY_FILE, index=False)

if not os.path.exists(INFO_FILE):
    info_df = pd.DataFrame(columns=["Name", "Email"])
    info_df.to_excel(INFO_FILE, index=False)

# Route to render survey form
@app.route('/')
def survey_form():
    return render_template('survey.html')

# Route to handle survey submission
@app.route('/submit-survey', methods=['POST'])
def submit_survey():
    responses = {
        'q1': request.form.get('q1'),
        'q2': request.form.get('q2'),
        'q3': request.form.get('q3'),
        'q4': request.form.get('q4'),
        'q5': request.form.get('q5'),
        'q6': request.form.get('q6'),
        'q7': request.form.get('q7'),
        'q8': request.form.get('q8'),
        'q9': request.form.get('q9'),
        'q10': request.form.get('q10'),
    }
    
    # Load the survey results
    df = pd.read_excel(SURVEY_FILE)
    
    # Update the tally for each response
    for question, answer in responses.items():
        new_entry = {'Question': question, 'Answer': answer}
        df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)

    # Save updated results
    df.to_excel(SURVEY_FILE, index=False)

    return render_template('name_email.html')

# Route to handle name and email submission
@app.route('/submit-info', methods=['POST'])
def submit_info():
    name = request.form.get('name')
    email = request.form.get('email')

    # Load user info sheet
    info_df = pd.read_excel(INFO_FILE)

    # Add the new entry
    new_entry = {'Name': name, 'Email': email}
    info_df = pd.concat([info_df, pd.DataFrame([new_entry])], ignore_index=True)

    # Save updated info
    info_df.to_excel(INFO_FILE, index=False)

    return "Thank you! Your information has been submitted."

# Start the Flask app
if __name__ == '__main__':
    app.run(debug=True)
