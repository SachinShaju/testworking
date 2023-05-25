from flask import Flask, render_template, request
from docx import Document
from docx.shared import Pt, RGBColor
from docx2pdf import convert
import yagmail

app = Flask(__name__)

@app.route('/')
def cert():
    return render_template('cert.html')

@app.route('/generate-certificate', methods=['POST'])

def generate_certificate():
    name = request.form['name']
    course = request.form['course']
    date = request.form['date']
    certificate_type = request.form['certificate_type']
    organization_name = request.form['organization_name']
    organizer_name = request.form['organizer_name']
    organizer_designation = request.form['organizer_designation'] 
    recipient_email = request.form['recipient_email']
    # Load template
    doc = Document('certificate_template.docx')

    # Replace fields in the template
    for p in doc.paragraphs:
        if 'Name' in p.text:
            p.text = p.text.replace('Name', name)
            p.runs[0].bold = True
            p.runs[0].font.size = Pt(32)  # Change the font size to 24
            p.runs[0].font.color.rgb = RGBColor(0x00,0x00,0x80)  # Set the font color to blue
        if 'CERTIFICATE_OF' in p.text:
            p.text = p.text.replace('CERTIFICATE_OF', certificate_type)
            p.runs[0].bold = True
            p.runs[0].font.size = Pt(35)  
            p.runs[0].font.color.rgb = RGBColor(0x00,0x00,0x80)
        if 'Course' in p.text:
            p.text = p.text.replace('Course', course)
            p.runs[0].font.size = Pt(26)
        if 'Date' in p.text:
            p.text = p.text.replace('Date', date)
            p.runs[0].font.size = Pt(26)
        if 'Organization_name' in p.text:
            p.text = p.text.replace('Organization_name', organization_name)
            p.runs[0].bold = True
            p.runs[0].font.size = Pt(28)  
        if 'Organizer_name' in p.text:
            p.text = p.text.replace('Organizer_name', organizer_name)
            p.runs[0].bold = True
            p.runs[0].font.size = Pt(24)
        if 'Organizer_designation' in p.text:
            p.text = p.text.replace('Organizer_designation', organizer_designation)
            p.runs[0].font.size = Pt(24)

    # Save the Word document 
    doc.save('certificate.docx')
    convert("certificate.docx")
    certificate_path = 'certificate.pdf'
    # Send the email with the certificate
    send_email_with_certificate(recipient_email, certificate_path)
    return certificate_path

def send_email_with_certificate(recipient_email, certificate_path):
    # Configure yagmail
    sender_email = 'certgenproject@gmail.com'  # Replace with your email address
    sender_password = 'bonqeueunpvqsvzz'  # Replace with your email password or app password
    subject = 'Certificate of Participation'

    # Create a yagmail object
    yag = yagmail.SMTP(sender_email, sender_password)

    # Send the email
    yag.send(
        to=recipient_email,
        subject=subject,
        contents='Please find the attached certificate.',
        attachments=[certificate_path]
    )
   
if __name__ == '__main__':
    app.debug = True
    app.run()

