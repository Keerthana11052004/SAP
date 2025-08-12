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
import pymysql

app = Flask(__name__, static_folder='templates/static')
app.config['SECRET_KEY'] = os.urandom(24)

# DB configuration
DB_HOST = 'localhost'
DB_USER = 'root'
DB_PASSWORD = 'Violin@12'
DB_NAME = 'approver_db'

# Email configuration
EMAIL_ADDRESS = os.environ.get('EMAIL_USER', 'sapnoreply@violintec.com')
EMAIL_PASSWORD = os.environ.get('EMAIL_PASS', 'VT$ofT@$2025')
SMTP_SERVER = os.environ.get('SMTP_SERVER', 'smtp.office365.com')
SMTP_PORT = os.environ.get('SMTP_PORT', 587)
RECIPIENT_EMAIL = 'automation1@violintec.com'

# Default cron schedule
# cron_schedule = os.environ.get('CRON_SCHEDULE', '*/30 * * * *')

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
    msg['Cc'] = 'narayanan.j@violintec.com'
    msg['X-Priority'] = '1 (Highest)'
    msg['X-MSMail-Priority'] = 'High'
    msg['Importance'] = 'High'

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
          <p>Dear {approver_name} ({recipient_email}),</p>
          <p>The following SAP documents are pending for your approval. Kindly review and approve them at the earliest:</p>
    """

    for doc_type, doc_numbers in grouped_data.items():
        html += f"""
            <p><strong>{doc_type}:</strong></p>
            <ul>
        """
        for doc_number in doc_numbers:
            html += f"<li>{doc_number}</li>"
        html += "</ul>"

    html += f"""
          <p><strong>SAP Link:</strong> https://my426081.s4hana.cloud.sap/ui#WorkflowTask-displayInbox</p>
          <br>
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
            connection = get_db_connection()
            print("DB connection established in fetch_and_send.")
            with connection.cursor() as cursor:
                cursor.execute("SELECT * FROM credentials ORDER BY id DESC LIMIT 1")
                credentials = cursor.fetchone()
                print(f"Credentials fetched from DB: {credentials}")
            connection.close()

            if credentials:
                api_url = credentials['api_url']
                username = credentials['username']
                password = credentials['password']
            else:
                api_url = None
                username = None
                password = None

            print(f"API URL: {api_url}, Username: {username}")

            if not all([api_url, username, password]):
                print("API credentials not found in database. Skipping fetch.")
                if not credentials:
                    print("Credentials object is None.")
                return

            auth_string = f"{username}:{password}"
            encoded_auth = base64.b64encode(auth_string.encode()).decode()
            auth_header = {"Authorization": f"Basic {encoded_auth}"}

            response = requests.get(api_url + "?$format=xml", headers=auth_header, verify=False)
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

def send_immediate_mail():
    """Fetches data and sends an email immediately."""
    print("Immediate mail sending process started.")
    fetch_and_send()
    print("Immediate mail sending process finished.")

@app.route('/', methods=['GET', 'POST'])
def index():
    print("Index route called!")
    data = None
    error = None
    success_message = None
    
    # Get credentials from request arguments
    api_url = request.args.get('odata_url', '')
    username = request.args.get('username', '')
    password = request.args.get('password', '')

    action = request.args.get('action')

    try:
        if api_url and username and password:
            # Save credentials to the database
            connection = get_db_connection()
            with connection.cursor() as cursor:
                cursor.execute("TRUNCATE TABLE credentials")
                print("Truncated credentials table.")
                cursor.execute("INSERT INTO credentials (api_url, username, password) VALUES (%s, %s, %s)", (api_url, username, password))
                print("Inserted new credentials.")
            connection.commit()
            connection.close()
            print("Saved credentials to database.")

            if action == 'fetch':
                auth_string = f"{username}:{password}"
                encoded_auth = base64.b64encode(auth_string.encode()).decode()
                auth_header = {"Authorization": f"Basic {encoded_auth}"}

                response = requests.get(api_url + "?$format=xml", headers=auth_header, verify=False)
                response.raise_for_status()

                root = ET.fromstring(response.content)
                entries = root.findall('.//{http://www.w3.org/2005/Atom}entry')
                data = []
                for entry in entries:
                    properties = {}
                    for prop in entry.findall('.//{http://schemas.microsoft.com/ado/2007/08/dataservices/metadata}properties/*'):
                        properties[prop.tag.split('}')[-1]] = prop.text
                    data.append(properties)
                success_message = "Data fetched successfully."

            elif action == 'send_mail':
                send_immediate_mail()
                success_message = "Immediate mail sent successfully."

    except requests.exceptions.RequestException as e:
        error = str(e)
    except Exception as e:
        error = str(e)

    # Fetch schedules to display on the page
    connection = get_db_connection()
    with connection.cursor() as cursor:
        cursor.execute("SELECT * FROM schedules ORDER BY hour, minute")
        schedules = cursor.fetchall()
    connection.close()

    return render_template('index.html', data=data, error=error, success_message=success_message, schedules=schedules)

@app.route('/add_schedule', methods=['POST'])
def add_schedule():
    minute = request.form.get('minute')
    hour = request.form.get('hour')
    day_of_month = request.form.get('day_of_month')
    month = request.form.get('month')
    day_of_week = request.form.get('day_of_week')
    api_url = request.form.get('odata_url')
    username = request.form.get('username')
    password = request.form.get('password')

    error = None
    success_message = None
    try:
        connection = get_db_connection()
        with connection.cursor() as cursor:
            cursor.execute("INSERT INTO schedules (minute, hour, day_of_month, month, day_of_week) VALUES (%s, %s, %s, %s, %s)", (minute, hour, day_of_month, month, day_of_week))
        connection.commit()
        connection.close()
        configure_scheduler() # Reconfigure scheduler with new schedule
        success_message = "Schedule added successfully."
    except Exception as e:
        error = f"Error adding schedule: {e}"

    return redirect(url_for('index', odata_url=api_url, username=username, password=password, success_message=success_message, error=error))

@app.route('/delete_schedule', methods=['POST'])
def delete_schedule():
    schedule_id = request.form.get('schedule_id')
    api_url = request.form.get('odata_url')
    username = request.form.get('username')
    password = request.form.get('password')

    error = None
    success_message = None
    try:
        connection = get_db_connection()
        with connection.cursor() as cursor:
            cursor.execute("DELETE FROM schedules WHERE id = %s", (schedule_id,))
        connection.commit()
        connection.close()
        configure_scheduler() # Reconfigure scheduler after deleting schedule
        success_message = "Schedule deleted successfully."
    except Exception as e:
        error = f"Error deleting schedule: {e}"

    return redirect(url_for('index', odata_url=api_url, username=username, password=password, success_message=success_message, error=error))
def get_db_connection():
    connection = pymysql.connect(host=DB_HOST,
                                 user=DB_USER,
                                 password=DB_PASSWORD,
                                 database=DB_NAME,
                                 cursorclass=pymysql.cursors.DictCursor)
    return connection

def init_db():
    connection = get_db_connection()
    with connection.cursor() as cursor:
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS schedules (
                id INT AUTO_INCREMENT PRIMARY KEY,
                minute VARCHAR(255) NOT NULL,
                hour VARCHAR(255) NOT NULL,
                day_of_month VARCHAR(255) NOT NULL,
                month VARCHAR(255) NOT NULL,
                day_of_week VARCHAR(255) NOT NULL
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS credentials (
                id INT AUTO_INCREMENT PRIMARY KEY,
                api_url VARCHAR(255) NOT NULL,
                username VARCHAR(255) NOT NULL,
                password VARCHAR(255) NOT NULL
            )
        """)
    connection.commit()
    connection.close()

def configure_scheduler():
    scheduler = BackgroundScheduler()
    try:
        connection = get_db_connection()
        with connection.cursor() as cursor:
            cursor.execute("SELECT * FROM schedules")
            schedules = cursor.fetchall()
        connection.close()

        if schedules:
            for schedule in schedules:
                scheduler.add_job(fetch_and_send, 'cron', minute=schedule['minute'], hour=schedule['hour'], day=schedule['day_of_month'], month=schedule['month'], day_of_week=schedule['day_of_week'])
        else:
            print("No schedules found in DB. No default cron job added.")
    except Exception as e:
        print(f"Error configuring scheduler from DB: {e}")

    if schedules:
        scheduler.start()

if __name__ == '__main__':
    init_db()
    configure_scheduler()
    app.run(debug=False)
