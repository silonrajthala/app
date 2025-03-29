The CreditCardSettEMAIL.exe application is designed to send Credit Card Statement emails. Please follow the instructions below to ensure proper functionality:

    1.SMTP Credentials: You will need to enter your SMTP credentials for smtp.office365.com.

    2.Directory Setup: Ensure that the directory E:\CreditCardStatement\Files exists on your computer. 
      This directory should contain files in the following format: (email;AcNo;date;accCode.pdf).

    3.Backup and Preparation:
        Before proceeding, back up all existing PDF files in the E:\CreditCardStatement\Files directory.
        After backing up, ensure that the E:\CreditCardStatement\Files directory is empty.

    4.File Transfer: Retrieve all PDF files from the following path: /online/mxpfo/Data/Revolving/Out/GIBL 
      and place them into the E:\CreditCardStatement\Files directory.

    5.RUN Appliction: Email Processing:
        Once the emails are sent, the corresponding files will be moved to E:\CreditCardStatement\FilesEmailSend.
        The success of the email delivery does not depend on whether the email address exists; if the email is sent successfully,
	the file will be moved regardless of the delivery status.

    6.Log File: After processing, please check E:\CreditCardStatement\log.txt for a detailed log of the email sending process.

    7.Skip List: The email addresses listed in E:\CreditCardStatement\Application\skipEmail.txt will be excluded from the email sending process, 
      even if there are corresponding files in the Files directory.

Thank you for using the CreditCardSettEMAIL application!