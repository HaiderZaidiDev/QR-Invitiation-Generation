import qrcode
import xlrd
from PIL import Image
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders


student_names = []
student_numbers = []
student_emails = []


# Opens workbook. 
wb = xlrd.open_workbook("excel-sheet.xlsx") 
sheet = wb.sheet_by_index(0) 

# Puts all student names in a list.
for i in range(6, 469, 1):
	student_names.append(sheet.cell_value(i, 1))  

# Puts all student numbers in a list
for k in range(6, 469, 1):
	student_numbers.append(int(sheet.cell_value(k, 2)))

#Saves QR codes as JPG's with each students full name and student numbers. 
for i in range(len(student_names)):
	img = qrcode.make('{0}, {1}'.format(student_numbers[i], student_names[i].replace(',', '')))
	img.save('filepath/{0}.jpg'.format(student_names[i])) # File saves as students last and first name. 

#Pastes QR codes onto the master copy of the ticket

image = Image.open('ticket-master.jpg')
position = (366, 93)
for j in range(len(student_names)):
	qrcodes = Image.open('filepath/{}.jpg'.format(student_names[j])).resize((129,129))
	image.paste(qrcodes, position)
	image.save('filepath/{}.jpg'.format(student_names[j]))

for h in range(len(student_numbers)):
	student_emails.append(str(student_numbers[h]) + '@districtdomain.net')


## sends email to each 12th grade student w/ qr code invitiation as attachment
email_user = 'email@gmail.com' # email sent from

for g in range(len(student_emails)):
	email_send = student_emails[g]
	subject = 'Invitiation (Test)'
	# Complex email as object

	msg = MIMEMultipart()
	msg['From'] = email_user
	msg['To'] = email_send
	msg['Subject'] = subject

	body = 'Hello, this is a diagnostic email for the JFSS Grad Luncheon (2019)'
	msg.attach(MIMEText(body, 'plain')) # Attaches body to email obj as plain txt

	## adding attachment to email
	filepath = 'filepath/{}.jpg'.format(student_names[g])
	filename = '{} Invitation.jpg'.format(student_names[g])
	attachment = open(filepath, 'rb')
	part = MIMEBase('application', 'octet-stream')
	part.set_payload(attachment.read())
	encoders.encode_base64(part)
	part.add_header('Content-Disposition', 'attachment;  filename={}'.format(filename))
	msg.attach(part)

	text = msg.as_string()
	# connects to gmail smptp with a.s.p
	server = smtplib.SMTP('smtp.gmail.com', 587)
	server.starttls() #establishes secure connection
	server.login(email_user, 'application_key' )

	server.sendmail(email_user, email_send, text) # Sends email to email_send with clean version of msg (text)
	print('Email sent.')
	server.quit()




