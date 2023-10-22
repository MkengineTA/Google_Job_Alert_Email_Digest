import pandas as pd
import imaplib
import email
import email.utils
from bs4 import BeautifulSoup
import re
from datetime import datetime

def extract_detailed_job_details(email_content, received_date):
    # Parsing the email's HTML body using BeautifulSoup
    soup = BeautifulSoup(email_content, 'html.parser')
    jobs_details = []
    job_elements = soup.find_all('a', href=True)

    for elem in job_elements:
        # Check for "+ 1 Filter" and skip if found
        if "+ 1 Filter" in elem.text:
            continue

        if "jobs" in elem.text:
            continue

        # Extracting job link
        job_link = elem["href"]

        # Extracting job title
        title_elem = elem.find('span', style=re.compile(r"font-size:.*?;"))
        title = title_elem.text if title_elem else None

        # Extracting company name
        company_elem = elem.find('div', style=re.compile(r"color: black;"))
        company = company_elem.text if company_elem else None

        # Extracting job location and keeping only the city
        # location_elem = elem.find('div', style=re.compile(r"color: #8A8A8A;"))
        # location = location_elem.text.split(",")[0] if location_elem else None
        location_elem = elem.find('div', style=re.compile(r"color: #8A8A8A;"))
        if location_elem:
            location_elems = location_elem.text.split(",")
            if len(location_elems) == 3:
                location = location_elems[1].strip()  # Take the city from "Postal Code, City, Country"
            elif len(location_elems) == 2:
                location = location_elems[0].strip()  # Take the city from "City, Country"
            else:
                location = None
        else:
            location = None

        # Extracting job date and type and adjusting the date format
        date_type_elems = elem.find_all('span', style=re.compile(r"color: #8A8A8A;"))
        date = None
        if date_type_elems:
            raw_date = date_type_elems[0].text
            date_match = re.match(r'(\d+\. \w+\.).*', raw_date)
            if date_match:
                day, month = date_match.group(1).split('. ')
                month_map = {
                    "Jan.": "01",
                    "Feb.": "02",
                    "März": "03",
                    "Apr.": "04",
                    "Mai": "05",
                    "Juni": "06",
                    "Juli": "07",
                    "Aug.": "08",
                    "Sep.": "09",
                    "Okt.": "10",
                    "Nov.": "11",
                    "Dez.": "12"
                }
                current_year = datetime.now().year
                current_month = datetime.now().month

                if month in month_map:
                    posting_month = int(month_map[month])

                    # Adjusting year based on the posting month and current month
                    if posting_month <= current_month:
                        # If posting month comes before the current month, assume it's from the current year
                        pass
                    else:
                        # Otherwise, assume it's from the previous year
                        current_year -= 1

                    date = f"{day.zfill(2)}.{month_map[month]}.{current_year}"
                else:
                    date = "keine Angabe"
        job_type = date_type_elems[1].text if len(date_type_elems) > 1 else None

        # Only appending if we have more than just the title and link
        if company or date or location or job_type:
            jobs_details.append((title, company, job_link, date, location, job_type, received_date))
    return jobs_details

def connect_to_gmail():
    print("Connecting to Gmail...")
    mail = imaplib.IMAP4_SSL('imap.gmail.com')
    mail.login('XXX@gmail.com', 'XXX')
    print("Connected to Gmail.")
    return mail

def fetch_emails(mail):
    mail.select('inbox')
    print("Fetching emails from jobalerts-noreply@google.com...")
    result, email_ids = mail.search(None, '(FROM "notify-noreply@google.com")')
    email_data = []
    for email_id in email_ids[0].split():
        result, email_content = mail.fetch(email_id, '(RFC822)')
        raw_email = email_content[0][1]
        email_data.append(email.message_from_bytes(raw_email))
    print(f"Fetched {len(email_data)} emails.")
    return email_data


def main():
    mail = connect_to_gmail()
    emails = fetch_emails(mail)
    all_jobs = []
    for msg in emails:
        received_date = email.utils.parsedate_to_datetime(msg['Date'])
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                if content_type == "text/html" and "attachment" not in content_disposition:
                    email_content = part.get_payload(decode=True).decode()
                    jobs = extract_detailed_job_details(email_content, received_date)
                    all_jobs.extend(jobs)
                    print(f"Extracted {len(jobs)} jobs from email.")
                    break
        else:
            email_content = msg.get_payload(decode=True).decode()
            jobs = extract_detailed_job_details(email_content, received_date)
            all_jobs.extend(jobs)
            print(f"Extracted {len(jobs)} jobs from email.")
    df = pd.DataFrame(all_jobs, columns=['Title', 'Company', 'Link', 'Date', 'Location', 'Job Type', 'Received Date'])
    df.sort_values(by='Received Date', inplace=True)
    df.drop_duplicates(subset=['Title', 'Location', 'Company'], keep='first', inplace=True)

    # Format the 'Received Date' column in "DD.MM.YYYY"
    df['Received Date'] = df['Received Date'].dt.strftime('%d.%m.%Y')

    # Fill empty Date fields with "keine Angabe"
    df['Date'].fillna("keine Angabe", inplace=True)

    df.to_excel("extracted_jobs.xlsx", index=False)
    print(f"Saved {len(df)} unique jobs to extracted_jobs.xlsx.")

# Adding the call to main() here
if __name__ == '__main__':
    main()