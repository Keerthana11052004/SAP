from flask import Flask, render_template, request, redirect, url_for
import requests
import xml.etree.ElementTree as ET
import base64
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from apscheduler.schedulers.background import BackgroundScheduler
from datetime import datetime

app = Flask(__name__, static_folder='templates/static')
app.config['SECRET_KEY'] = os.urandom(24)

# Email configuration
EMAIL_ADDRESS = os.environ.get('EMAIL_USER', 'sapnoreply@violintec.com')
EMAIL_PASSWORD = os.environ.get('EMAIL_PASS', 'VT$ofT@$2025')
SMTP_SERVER = os.environ.get('SMTP_SERVER', 'smtp.office365.com')
SMTP_PORT = os.environ.get('SMTP_PORT', 587)
RECIPIENT_EMAIL = 'automation1@violintec.com'

# Default cron schedule
cron_schedule = os.environ.get('CRON_SCHEDULE', '*/30 * * * *')

def send_email(recipient_email, data):
    # Group documents by SAPObjectNodeRepresentation
    grouped_data = {}
    for item in data:
        doc_type = item.get('SAPObjectNodeRepresentation')
        if doc_type not in grouped_data:
            grouped_data[doc_type] = []
        grouped_data[doc_type].append(item.get('SAPBusinessObjectNodeKey1'))

    msg = MIMEMultipart('alternative')
    msg['Subject'] = "SAP documents pending for your approval"
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = RECIPIENT_EMAIL

    approver_name = "Approver"
    if data:
        first_name = data[0].get('FirstName', '')
        last_name = data[0].get('LastName', '')
        if first_name or last_name:
            approver_name = f"{first_name} {last_name}".strip()

    html = f"""
    <html>
      <head></head>
      <body>
        <div>
          <p>Dear {approver_name},</p>
          <p>The following SAP documents are pending for your approval. Kindly review and approve them at the earliest:</p>
    """

    for doc_type, doc_numbers in grouped_data.items():
        html += f"""
            <details>
                <summary>{doc_type}</summary>
        """
        for doc_number in doc_numbers:
            html += f"""
                <div class="doc-number">{doc_number}</div>
            """
        html += f"""
            </details>
        """

    html += f"""
          <p>Regards,<br>
           Enterprise Automation Office</p>
        </div>
      </body>
    </html>
    """

    part2 = MIMEText(html, 'html')
    msg.attach(part2)

    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        server.sendmail(EMAIL_ADDRESS, RECIPIENT_EMAIL, msg.as_string())
        server.quit()
        print(f"Email sent successfully to {RECIPIENT_EMAIL} at {datetime.now()}")
    except Exception as e:
        print(f"Failed to send email: {e}")

def fetch_and_send():
    with app.app_context():
        data = None
        error = None
        print(f"fetch_and_send called at {datetime.now()}")
        try:
            api_url = os.environ.get('API_URL')
            username = os.environ.get('API_USERNAME')
            password = os.environ.get('API_PASSWORD')
            print(f"API URL: {api_url}, Username: {username}")

            auth_string = f"{username}:{password}"
            encoded_auth = base64.b64encode(auth_string.encode()).decode()
            auth_header = {"Authorization": f"Basic {encoded_auth}"}

            response = requests.get(api_url + "?$format=xml", headers=auth_header)
            response.raise_for_status()

            root = ET.fromstring(response.content)
            entries = root.findall('.//{http://www.w3.org/2005/Atom}entry')
            all_data = []
            for entry in entries:
                properties = {}
                for prop in entry.findall('.//{http://schemas.microsoft.com/ado/2007/08/dataservices/metadata}properties/*'):
                    properties[prop.tag.split('}')[-1]] = prop.text
                all_data.append(properties)

            if all_data:
                print(f"Data fetched: {len(all_data)} items. Grouping by email and sending...")
                # Group data by WorkplaceAddress (email)
                grouped_by_email = {}
                for item in all_data:
                    email = item.get('EmailAddress')
                    if email:
                        if email not in grouped_by_email:
                            grouped_by_email[email] = []
                        grouped_by_email[email].append(item)

                # Send email to each user
                for email, user_data in grouped_by_email.items():
                    send_email(email, user_data)
            else:
                print("No data fetched, not sending email.")
        except requests.exceptions.RequestException as e:
            error = str(e)
            print(f"Error fetching data: {error}")
        except Exception as e:
            error = str(e)
            print(f"Error: {error}")

@app.route('/', methods=['GET', 'POST'])
def index():
    print("Index route called!")
    data = None
    error = None
    success_message = None
    try:
        api_url = request.args.get('odata_url', '')
        username = request.args.get('username', '')
        password = request.args.get('password', '')

        os.environ['API_URL'] = api_url
        os.environ['API_USERNAME'] = username
        os.environ['API_PASSWORD'] = password

        auth_string = f"{username}:{password}"
        encoded_auth = base64.b64encode(auth_string.encode()).decode()
        auth_header = {"Authorization": f"Basic {encoded_auth}"}

        response = requests.get(api_url + "?$format=xml", headers=auth_header)
        response.raise_for_status()

        root = ET.fromstring(response.content)
        entries = root.findall('.//{http://www.w3.org/2005/Atom}entry')
        data = []
        for entry in entries:
            properties = {}
            for prop in entry.findall('.//{http://schemas.microsoft.com/ado/2007/08/dataservices/metadata}properties/*'):
                properties[prop.tag.split('}')[-1]] = prop.text
            data.append(properties)

        if data:
            # Group data by WorkplaceAddress (email)
            grouped_by_email = {}
            for item in data:
                email = item.get('EmailAddress')
                if email:
                    if email not in grouped_by_email:
                        grouped_by_email[email] = []
                    grouped_by_email[email].append(item)

            # Send email to each user
            for email, user_data in grouped_by_email.items():
                send_email(email, user_data)

    except requests.exceptions.RequestException as e:
        error = str(e)
    except Exception as e:
        error = str(e)

    return render_template('index.html', data=data, error=error, success_message=success_message, cron_expression=cron_schedule)

@app.route('/schedule', methods=['POST'])
def schedule():
    minute = request.form.get('minute')
    hour = request.form.get('hour')
    day_of_month = request.form.get('day_of_month')
    month = request.form.get('month')
    day_of_week = request.form.get('day_of_week')
    api_url = request.form.get('odata_url')
    username = request.form.get('username')
    password = request.form.get('password')

    success_message = None
    error = None

    try:
        os.environ['API_URL'] = api_url
        os.environ['API_USERNAME'] = username
        os.environ['API_PASSWORD'] = password

        cron_expression = f"{minute} {hour} {day_of_month} {month} {day_of_week}"

        os.environ['CRON_SCHEDULE'] = cron_expression
        print(f"Schedule updated to: {cron_expression}")
        # Removed test email sending, as user wants to test with actual data
        # test_data = [{'SAPObjectNodeRepresentation': 'Test', 'SAPBusinessObjectNodeKey1': '123'}]
        # send_email(test_data)
        success_message = "Schedule updated. Please set your API details and fetch data to test."
    except Exception as e:
        print(f"Invalid CRON expression: {e}")
        error = f"Invalid CRON expression: {e}"
    return redirect(url_for('index', odata_url=api_url, username=username, password=password, success_message=success_message, error=error))

def configure_scheduler():
    scheduler = BackgroundScheduler()
    # Parse the cron_schedule string into components
    try:
        minute, hour, day_of_month, month, day_of_week = cron_schedule.split()
        scheduler.add_job(fetch_and_send, 'cron', minute=minute, hour=hour, day=day_of_month, month=month, day_of_week=day_of_week)
    except ValueError:
        # Fallback to default if parsing fails
        scheduler.add_job(fetch_and_send, 'cron', minute='*/30')
    scheduler.start()

if __name__ == '__main__':
    configure_scheduler()
    app.run(debug=True)
