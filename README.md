# QR Invitiation Generation
A QR invitiation generation script developed for John Fraser Secondary School's 2019 Grad Luncheon. 

## Features
* Generates personalized invitiations for all graduating students, extracting the required information (e.g full name and student number) from a spreadsheet provided by the school's administraiton. 
* Emails invitiations to each student's district issued email address. 

Each invitiation contains:
* The respective student's full name and homeroom class, allowing for easy distribution. 
* A QR code, containing the students full name and student number to be scanned at the entry point of the venue. 

## Invitiation Verificiation & Tracking
To verify invitiations, the QR codes on all invitiations were scanned via "Scan To Sheets", a free application available on the iOS app store. This application uploaded the data from the scanned QR codes onto a Google Spreadsheet. The invitiations were validated using conditonal formatting, scanning against a master column. Where valid tickets were highlighted green, and invalid red.
