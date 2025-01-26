import configparser
import sys
from os import PathLike

from docx import Document as DocxDocument
from datetime import datetime, timedelta
import os
import subprocess
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders


def replace_text_in_docx(doc_path, replacements):
    doc = DocxDocument(doc_path)

    for paragraph in doc.paragraphs:
        for old_text, new_text in replacements.items():
            if old_text in paragraph.text:
                paragraph.text = paragraph.text.replace(old_text, new_text)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for old_text, new_text in replacements.items():
                        if old_text in paragraph.text:
                            paragraph.text = paragraph.text.replace(old_text, new_text)

    return doc


def convert_to_pdf(docx_path, pdf_path):
    libreoffice_path = "/usr/bin/libreoffice"  # This is the default path on Ubuntu
    if not os.path.exists(libreoffice_path):
        libreoffice_path = "/usr/bin/soffice"  # Try alternate path

    if not os.path.exists(libreoffice_path):
        raise FileNotFoundError("LibreOffice not found. Please install it.")

    subprocess.run([
        libreoffice_path,
        '--headless',
        '--convert-to', 'pdf',
        '--outdir', os.path.dirname(pdf_path),
        docx_path
    ], check=True)

    # Rename the output file to match the desired pdf_path
    os.rename(os.path.splitext(docx_path)[0] + '.pdf', pdf_path)


def make_quittance(property_dict: dict) -> str:
    address = dict(property_dict)["address"]
    tenantname = dict(property_dict)["tenantname"]
    landlordname = dict(property_dict)["landlordname"]
    hors_charge = dict(property_dict)["hors_charge"]
    charge = dict(property_dict)["charge"]
    total_litteral = dict(property_dict)["total_litteral"]
    rent_amt = int(hors_charge) + int(charge)
    date_format = "%d-%m-%Y"

    # Calculate dates
    months_offset = 1
    today = datetime.today().replace(month=datetime.today().month - months_offset)  # Replace with today's date
    first_day_this_month = today.replace(day=1).strftime(date_format)
    first_day_next_month = (today.replace(day=1) + timedelta(days=32)).replace(day=1).strftime(date_format)
    thirteenth_this_month = today.replace(day=13).strftime(date_format)
    twenty_fifth_this_month = today.replace(day=25).strftime(date_format)

    # Define replacements
    replacements = {
        "{address}": address,
        "{tenantname}": tenantname,
        "{rent_letters}": total_litteral,
        "{rent_amt}": str(rent_amt) + "€",
        "{start}": first_day_this_month,
        "{end}": first_day_next_month,
        "{ex_charge}": str(hors_charge) + "€",
        "{charge}": str(charge) + "€",
        "{recu}": thirteenth_this_month,
        "{signed}": twenty_fifth_this_month
    }
    output_doc = f'{property_name}_quittance_{today.strftime(date_format)}.docx'
    output_pdf = f'{property_name}_quittance_{today.strftime(date_format)}.pdf'
    modified_doc = replace_text_in_docx(input_doc, replacements)
    modified_doc.save(output_doc)
    print(f"Word document has been updated and saved as {output_doc}")
    # Convert to PDF
    convert_to_pdf(output_doc, output_pdf)
    print(f"PDF has been created as {output_pdf}")
    os.remove(output_doc)

    msg_body = f"Cher {tenantname},\n\nVeuillez trouver ci-joint votre quittance de loyer du {first_day_this_month} au {first_day_next_month}.\n\nMerci,\n{landlordname}"

    return msg_body, output_pdf

def validate_email_params(sender_email, sender_password, receiver_email):
    """Validate email parameters before sending."""
    if not all([sender_email, sender_password, receiver_email]):
        raise ValueError("Missing required email parameters")
    if '@' not in sender_email or '@' not in receiver_email:
        raise ValueError("Invalid email format")

def send_email(sender_email, sender_password, receiver_email, subject, body, attachment_path):
    """Send email with improved error handling and validation."""
    try:
        validate_email_params(sender_email, sender_password, receiver_email)
        
        if not os.path.exists(attachment_path):
            raise FileNotFoundError(f"Attachment file not found: {attachment_path}")

        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Subject'] = subject

        msg.attach(MIMEText(body, 'plain'))

        # Handle attachment
        with open(attachment_path, "rb") as attachment:
            part = MIMEBase("application", "pdf")  # Specify PDF type
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename= {os.path.basename(attachment_path)}",
            )
            msg.attach(part)

        # Connect to SMTP server with proper error handling
        server = None
        try:
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(sender_email, sender_password)
            text = msg.as_string()
            server.sendmail(sender_email, receiver_email, text)
        finally:
            if server:
                server.quit()

        # Clean up the PDF file after successful sending
        os.remove(attachment_path)
        return True

    except smtplib.SMTPAuthenticationError:
        raise Exception("Failed to authenticate with Gmail. Please check your credentials and ensure you're using an App Password.")
    except smtplib.SMTPException as e:
        raise Exception(f"SMTP error occurred: {str(e)}")
    except Exception as e:
        raise Exception(f"Failed to send email: {str(e)}")


def validate_config(config):
    """Validate the configuration file."""
    if 'gmail' not in config.sections():
        raise ValueError("Missing 'gmail' section in config file")
    
    gmail_config = dict(config.items('gmail'))
    if not gmail_config.get('sender_email') or not gmail_config.get('sender_password'):
        raise ValueError("Missing email credentials in config file")

    properties = [s for s in config.sections() if s != 'gmail']
    if not properties:
        raise ValueError("No property sections found in config file")

    required_fields = ['address', 'tenantname', 'tenant_email', 'landlordname', 'hors_charge', 'charge', 'total_litteral']
    for prop in properties:
        prop_dict = dict(config.items(prop))
        missing = [f for f in required_fields if not prop_dict.get(f)]
        if missing:
            raise ValueError(f"Missing required fields {missing} in property section {prop}")

def main():
    """Main function with improved error handling."""
    try:
        input_doc = "quittance_template.docx"
        if not os.path.exists(input_doc):
            raise FileNotFoundError(f"Template file not found: {input_doc}")

        config_path = os.path.join(os.sep, os.getcwd(), 'config.ini')
        if not os.path.exists(config_path):
            raise FileNotFoundError(f"Config file not found: {config_path}")

        ini_file = configparser.ConfigParser(interpolation=None)
        ini_file.read(config_path)
        
        # Validate config file
        validate_config(ini_file)

        # Get email credentials
        sender_email = dict(ini_file.items(section="gmail"))["sender_email"]
        sender_password = dict(ini_file.items(section="gmail"))["sender_password"]
        
        # Process each property
        properties_list = [s for s in ini_file.sections() if s != 'gmail']
        for property_name in properties_list:
            try:
                property_dict = dict(ini_file.items(section=property_name))
                msg_body, output_pdf = make_quittance(property_dict)

                receiver_email = property_dict['tenant_email']
                subject = f"Quittance de loyer {property_name}"

                if send_email(sender_email, sender_password, receiver_email, subject, msg_body, output_pdf):
                    print(f"Email sent successfully for property: {property_name}")
            except Exception as e:
                print(f"Failed to process property {property_name}: {str(e)}")
                # Continue with next property even if one fails
                continue

    except Exception as e:
        print(f"Error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    import sys
    main()
