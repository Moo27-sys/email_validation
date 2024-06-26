from flask import Flask, request, render_template, send_file
import requests
import logging
from logging.handlers import RotatingFileHandler
import os
import pandas as pd
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

app = Flask(__name__, template_folder='template')
app.secret_key = 'your_secret_key'

API_KEY = "Btc7U5ZLK63fCimLj61MbhoOq7YREDpF"
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
SENDER_EMAIL = 'monkhm22@gmail.com'
RECEIVER_EMAIL = 'monkhm02@gmail.com'
APP_SPECIFIC_PASSWORD = 'gqig cwty xopb vbyi'

# Configure logging
if not os.path.exists('logs'):
    os.mkdir('logs')
file_handler = RotatingFileHandler('logs/email_checker.log', maxBytes=10240, backupCount=10)
file_handler.setFormatter(logging.Formatter(
    '%(asctime)s %(levelname)s: %(message)s [in %(pathname)s:%(lineno)d]'
))
file_handler.setLevel(logging.INFO)
app.logger.addHandler(file_handler)
app.logger.setLevel(logging.INFO)
app.logger.info('Email Checker startup')

class Validate:
    def __init__(self, key):
        self.key = key
        self.base_url = f"https://ipqualityscore.com/api/json/email/{self.key}/"

    def email_validation_api(self, email: str) -> dict:
        params = {
            "email": email
        }
        try:
            response = requests.get(self.base_url, params=params)
            response.raise_for_status()
            app.logger.info(f"API response: {response.text}")
            return response.json()
        except requests.exceptions.HTTPError as e:
            app.logger.error(f"HTTP error occurred: {str(e)}")
            raise
        except requests.exceptions.RequestException as e:
            app.logger.error(f"Request error occurred: {str(e)}")
            raise
        except ValueError as e:
            app.logger.error(f"Invalid JSON response: {str(e)}")
            raise

    def is_suspicious(self, result: dict) -> bool:
        return (
            result.get('disposable') or
            result.get('spam_trap') or
            result.get('recent_abuse') or
            result.get('fraud_score', 0) > 75
        )

def send_suspicious_email(subject: str, body: str):
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = RECEIVER_EMAIL
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SENDER_EMAIL, APP_SPECIFIC_PASSWORD)
        text = msg.as_string()
        server.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, text)
        server.quit()
        app.logger.info(f"Suspicious email sent: {subject}")
    except Exception as e:
        app.logger.error(f"Failed to send email: {str(e)}")

@app.errorhandler(Exception)
def handle_exception(e):
    app.logger.error(f"An error occurred: {str(e)}")
    return render_template('error.html', error=str(e)), 500

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        email = request.form['email']
        validator = Validate(API_KEY)
        try:
            result = validator.email_validation_api(email)
            app.logger.info(f"Validation result for {email}: {result}")
            if validator.is_suspicious(result):
                send_suspicious_email(f'Suspicious Email Detected: {email}', str(result))
            return render_template('result.html', result=result, email=email)
        except requests.RequestException as e:
            app.logger.error(f"API request failed: {str(e)}")
            return render_template('error.html', error="Failed to contact the email validation service. Please try again later.")
        except Exception as e:
            app.logger.error(f"An unexpected error occurred: {str(e)}")
            return render_template('error.html', error="An unexpected error occurred. Please try again.")
    return render_template('index.html')

@app.route('/bulk', methods=['GET', 'POST'])
def bulk():
    if request.method == 'POST':
        if 'file' not in request.files:
            return render_template('error.html', error="No file part")
        file = request.files['file']
        if file.filename == '':
            return render_template('error.html', error="No selected file")
        if file and (file.filename.endswith('.csv') or file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
            try:
                # Read the file
                if file.filename.endswith('.csv'):
                    df = pd.read_csv(file)
                else:
                    df = pd.read_excel(file)
                
                app.logger.info(f"File read successfully. Shape: {df.shape}")
                
                validator = Validate(API_KEY)
                results = []
                suspicious_emails = []

                for index, email in enumerate(df.iloc[:, 0]):
                    try:
                        app.logger.info(f"Processing email {index + 1}: {email}")
                        result = validator.email_validation_api(email)
                        result['email'] = email  # Include the email in the result
                        result['suspicious'] = validator.is_suspicious(result)  # Add suspicious flag
                        results.append(result)
                        if result['suspicious']:
                            suspicious_emails.append((email, result))
                    except Exception as e:
                        app.logger.error(f"Error validating {email}: {str(e)}")
                        results.append({
                            'email': email,
                            'Status': 'Error',
                            'Sub_Status': str(e),
                            'suspicious': True
                        })

                app.logger.info(f"Processed {len(results)} emails")

                # Create a DataFrame from the results
                result_df = pd.DataFrame(results)

                # Reorder columns to make sure 'email' is the first column
                columns = ['email'] + [col for col in result_df.columns if col != 'email']
                result_df = result_df[columns]

                # Create an Excel file
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False, sheet_name='Validation Results')

                output.seek(0)
                app.logger.info("Excel file created successfully")

                # Send suspicious emails if any
                if suspicious_emails:
                    email_body = "\n\n".join([f"{email}\n{details}" for email, details in suspicious_emails])
                    send_suspicious_email("Bulk Processed Suspicious Emails", email_body)

                return send_file(
                    output,
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    as_attachment=True,
                    download_name='validation_results.xlsx'
                )
            except Exception as e:
                app.logger.error(f"An error occurred during bulk processing: {str(e)}")
                return render_template('error.html', error=f"An error occurred during bulk processing: {str(e)}")
        else:
            return render_template('error.html', error="Invalid file format. Please upload a CSV or Excel file.")
    return render_template('bulk.html')

if __name__ == "__main__":
    app.run(debug=True, host='localhost', port=8000)
