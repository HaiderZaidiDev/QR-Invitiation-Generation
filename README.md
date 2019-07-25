# QR Invitiation Generation
A QR invitiation generation script developed for John Fraser Secondary School's 2019 Grad Luncheon. 

## Features
* Generates personalized invitiations for all graduating students, extracting the required information (e.g full name and student number) from a spreadsheet provided by the school's administraiton. 
* Emails invitiations to each student's district issued email address. 

Each invitiation contains:
* The respective student's full name and homeroom class, allowing for easy distribution. 
* A QR code, containing the students full name and student number to be scanned at the entry point of the venue. 

## Invitiation Verificiation & Tracking
To verify invitiations, the QR codes on all invitiations were scanned via "Scan To Sheets", a free application available on the iOS app store. This application inputted the data from the QR codes onto a google spreadsheet. The data from the QR codes were then verified using conditional formatting by checking if the data was found in a column containing the data from all the QR codes generated. If the invitiation was validated succesfully, the cell containing the scanned data would highlight green notifying the person scanning invitiations to allow admission into the venue and stamp their hands (such that students could leave the venue and come back if they wish without having to scan again). If the invitiation was not valid, the cell would highlight red, and if the invitiation had been scanned more then once the cell would highlight yellow (in which case the person scanning would check to see if the student has been stamped).
