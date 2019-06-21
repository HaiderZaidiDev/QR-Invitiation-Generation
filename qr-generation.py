import qrcode
import xlrd
from PIL import Image, ImageDraw, ImageFont
import xlwt
import time
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders



student_names = []
student_numbers = []
sem2_hr = []
student_emails = []

fileprefix = '' # File path of folder containing the script and spreadsheet of students. 
font = ImageFont.truetype('Roboto-Bold.ttf', size=45)

# Opens spreadsheet with RSVP'd student attendees. 
wb_one = xlrd.open_workbook("attendees.xlsx") 
attendees = wb_one.sheet_by_index(0)

# Puts all student names in a list.
for i in range(0, 367, 1):
	student_names.append(attendees.cell_value(i, 1))  

# Puts all student numbers in a list.
for j in range(0, 367, 1):
	student_numbers.append(int(attendees.cell_value(j, 2)))

# Puts each student's semester two homeroom in a list.
for k in range(0,367, 1):
	sem2_hr.append(attendees.cell_value(k, 5))


# Saves QR codes as JPG's with each students full name and student numbers. 
for i in range(len(student_names)):
	img = qrcode.make('{0}, {1}'.format(student_numbers[i], student_names[i].replace(',', ''))) 
	img.save('{0}/codes/{1}.jpg'.format(fileprefix, student_names[i])) # Saves image containing QR code, name is formatted with the student's last then first name. 



# Pasting each QR code onto invitiations at designated cordinates. 
for j in range(len(student_names)):
	image = Image.open('invitation-master.png')
	qrcodes = Image.open('{0}/codes/{1}.jpg'.format(fileprefix, student_names[j])).resize((410,410))
	image.paste(qrcodes, position)

	draw = ImageDraw.Draw(image)
 
	namepos = (2178, 642)
	color = 'rgb(255, 255, 255)' # black color
	draw.text(namepos, student_names[j], fill=color, font=font)

	hrpos = (2280,700)
	draw.text(hrpos, sem2_hr[0], fill=color, font=font)

	image.save('{0}/tickets/{1}.png'.format(fileprefix, student_names[j]))
	image.close()


for h in range(len(student_numbers)):
	student_emails.append(str(student_numbers[h]) + '@districtdomain.net') # Creates a list of student emails deriving from student numbers. 

print('Ticket compiling complete.')


email_user = 'mystudentnumber@districtdomain.net' 

# Sends each invitiation to the corresponding student via email. 
for g in range(len(student_emails)):
	email_send = student_emails[g]
	subject = 'JFSS Grad Luncheon (2019) Invitiation'
	# Complex email as object

	msg = MIMEMultipart()
	msg['From'] = email_user
	msg['To'] = email_send
	msg['Subject'] = subject

	body = 'Hello, attached below is your digital personal invitation to the JFSS Grad Luncheon being held on Thursday, June 6th in the Cafeteria at the beginning of lunch. You should have been given hard-copies of your invitiation by your homeroom teacher, if you have not (or if you lost it), you can present this ticket at event entry. Please note you are REQUIRED to present your ticket to enter the cafeteria for the grad luncheon, whether it be a digital or hard-copy version. If you have any questions or concerns, please don’t hesitate to respond to this email or reach out to us via Instagram (@jfssluncheon2019). '
	msg.attach(MIMEText(body, 'plain')) # Attaches body to email object as plain text.

	# Adds the invitiations as attachments to the email, attachment name formatted as "{Last Name}, {FirstName} Invitiation".
	filepath = '{0}/tickets/{1}.png'.format(fileprefix, student_names[g])
	filename = '{} Invitation.png'.format(student_names[g])
	attachment = open(filepath, 'rb')
	part = MIMEBase('application', 'octet-stream')
	part.set_payload(attachment.read())
	encoders.encode_base64(part)
	part.add_header('Content-Disposition', 'attachment;  filename={}'.format(filename))
	msg.attach(part)

	text = msg.as_string()
	# Connects to gmail smtp with application specific password. 
	server = smtplib.SMTP('smtp.gmail.com', 587)
	server.starttls() #establishes secure connection
	server.login(email_user, 'APP_SPECIFIC_PASS')

	server.sendmail(email_user, email_send, text) # Sends email to email_send with clean version of msg (text)
	print('Email sent.')
	server.quit()
