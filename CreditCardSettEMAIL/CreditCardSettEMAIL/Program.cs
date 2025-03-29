using System;
using System.IO; // For StreamWriter
using System.Net;
using System.Net.Mail;
using System.Collections.Generic;

using System.Threading.Tasks;

using System.Text.RegularExpressions;

using System.Linq; // Required for Contains

namespace CreditCardSettEMAIL
{
    class Program
    {
        public class Recipient
        {
            public string Email { get; set; }
            public string AccountNumber { get; set; }
            public string Attachment { get; set; }
            public string Date { get; set; }
            public string CardCode { get; set; }

            public Recipient(string email, string accountnumber, string attachment, string date, string cardcode)
            //public Recipient(string email, string accountnumber, string attachment, string date)
            {
                Email = email;
                AccountNumber = accountnumber;
                Attachment = attachment;
                Date = date;
                CardCode = cardcode;
            }
        }

        static void Main(string[] args)
        {
            string logFilePath = @"E:\CreditCardStatement\log.txt"; // Update with your desired log file path
            string pdfDirectory = @"E:\CreditCardStatement\Files"; // Directory containing PDF files
            string destinationDirectory = @"E:\CreditCardStatement\FilesEmailSend"; // this path contain file that email sent sucessfully
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12; // Set the security protocol to use TLS 1.2

            using (StreamWriter logWriter = new StreamWriter(logFilePath, true)) // Append mode
            {
                logWriter.AutoFlush = true; // Ensure that the log is flushed immediately
                Console.WriteLine("Processing to Send Mail!");

                logWriter.WriteLine(string.Format("\nRunning Send Mail! Application On {0}", DateTime.Now));
                logWriter.WriteLine(string.Format("Processing to Send Mail!"));


                while (true) // Loop until valid input is received
                {

                    Console.Write("Do You need Instructions (y/n) ?:");
                    string instruction = Console.ReadLine().ToLower(); // Read input and convert to lowercase
                    if (instruction == "y")
                    {
                        // Display instructions from instru.txt
                        try
                        {
                            string instructions = File.ReadAllText(@"E:\CreditCardStatement\Application\README.txt");
                            Console.WriteLine("Instructions:\n");
                            Console.WriteLine(instructions);
                        }
                        catch (FileNotFoundException)
                        {
                            Console.WriteLine("Instructions file not found.");
                        }
                        break; // Exit the loop after displaying instructions
                    }
                    else if (instruction == "n")
                    {
                        Console.WriteLine("Continuing without instruction. Guessing you have Read README.txt OR You are familiar with this application...");
                        break; // Exit the loop to continue with the program
                    }
                    else
                    {
                        Console.WriteLine("Invalid input. Please enter 'y' or 'n'.");
                    }
                }

                // Prompt user for email credentials
                Console.Write("\n\nSMTP server:smtp.office365.com\n");

                Console.Write("Enter the email address: ");
                string userEmail = Console.ReadLine();

                Console.Write("Enter the email password: ");
                string userPassword = ReadPassword(); // Custom method to read password without echoing

                // Set up the SMTP client
                SmtpClient smtpClient = new SmtpClient("smtp.office365.com")
                {
                    Port = 587,
                    Credentials = new NetworkCredential(userEmail, userPassword), // Use user-provided credentials
                    EnableSsl = true,
                };

                // List to hold recipients
                var recipients = new List<Recipient>();

                // Read PDF files from the specified directory
                foreach (var file in Directory.GetFiles(pdfDirectory, "*.pdf"))
                {
                    // Extract email and account number from the filename
                    string fileName = Path.GetFileNameWithoutExtension(file); // Get the file name without extension
                    string[] parts = fileName.Split(';'); // Split by semicolon

                    if (parts.Length == 4)
                   //     if (parts.Length == 3)
                    {
                        string email = parts[0];
                        string accountNumber = parts[1];
                        string date = parts[2];
                        string accountCode = parts[3];

                        // Validate email format
                        if (IsValidEmail(email))
                        {
                            // Create a Recipient object and add it to the list
                            recipients.Add(new Recipient(email, accountNumber, file, date, accountCode));
                          //  recipients.Add(new Recipient(email, accountNumber, file, date));
                        }
                        else
                        {
                            Console.WriteLine(string.Format("Invalid email format: '{0}' in filename '{1}'.",email,fileName));
                            logWriter.WriteLine(string.Format("ERROR {0}: Invalid email format: '{1}' in filename '{2}.pdf'", DateTime.Now, email, fileName));
                        }
                    }
                    else
                    {
                        Console.WriteLine(string.Format("Filename '{0}' does not match expected format.",fileName));
                        logWriter.WriteLine(string.Format("WARNING {0}: Filename '{1}' does not match expected format.",DateTime.Now,fileName));
                    }
                }
                // Read emails to skip from a text file
                string[] skipEmails = File.ReadAllLines( @"E:\CreditCardStatement\Application\skipEmail.txt");

                // Continue with the rest of your email sending logic...
                foreach (var recipient in recipients)
                {
                    if (skipEmails.Contains(recipient.Email))
                        {
                            //Console.WriteLine($"Skipping email to {recipient.Email} as it is in the skip list.");
                            Console.WriteLine(string.Format("Skipping email to {0} as it is in the skip list.", recipient.Email));

                            //logWriter.WriteLine($"{DateTime.Now}: Skipping email to {recipient.Email} as it is in the skip list.");
                            logWriter.WriteLine(string.Format("{0}: Skipping email to {1} as it is in the skip list.", DateTime.Now,recipient.Email));

                            continue; // Skip to the next recipient
                        }
                    // Attach the file if it exists
                    if (File.Exists(recipient.Attachment))
                    {                       
                        string date = string.Format("{0}", recipient.Date); // Assuming recipient.Date is in "dd-MM-yyyy" format

                        DateTime parsedDate;
                        string formattedDate = string.Empty; // Declare formattedDate outside the if block
                        bool emailStatus = false; // Declare formattedDate outside the if block

                        // Parse the date string using the specified format
                        if (DateTime.TryParseExact(date, "dd-MM-yyyy", null, System.Globalization.DateTimeStyles.None, out parsedDate))
                        {
                            // If the date is valid, calculate the previous month and year
                            int previousMonth = parsedDate.Month - 1;
                            int previousYear = parsedDate.Year;

                            // If the previous month is less than 1, we need to go to December of the previous year
                            if (previousMonth < 1)
                            {
                                previousMonth = 12;
                                previousYear--;
                            }

                            // Create a DateTime for the previous month
                            DateTime previousDate = new DateTime(previousYear, previousMonth, 1);

                            // Format the previous month to "MMMM, yyyy"
                            formattedDate = previousDate.ToString("MMMM, yyyy");
                        }
                        else
                        {
                            Console.WriteLine("Invalid date or Date not found.");
                            logWriter.WriteLine(string.Format("Invalid date for email {0}", recipient.Email));
                        }
                        if (string.IsNullOrEmpty(formattedDate))
                        {
                            formattedDate = "NULL"; // Exit the method or handle accordingly
                        }

                        // Create the MailMessage object
                        MailMessage mailMessage = new MailMessage
                        {
                            
                           // From = new MailAddress("alert@sct.com.np"), // Replace with your email
                           // From = new MailAddress("statement@sct.com.np"), // Replace with your email
                            From = new MailAddress(userEmail), // Replace with your email
                            Subject = "Credit Card Statement",
                            Body = string.Format(
                                "<html><body>" +
                                "<p>Dear Valued Customer,</p>" +
                                "<p>Please find your attached <b>Visa Credit Card</b> statement for the month of <b>{0}</b>. We kindly request you to review the Transaction details enlisted and ensure the timely payment of any outstanding dues within due date to avoid interest, late fees, or additional charges.</p>" +
                                "<p>Should you notice any discrepancies or require clarification regarding the statement, please do not hesitate to contact us. Our team will be happy to assist you.</p>" +
                                 "<p>Please feel free to reach us at: <br><b>Email:</b> <a href='mailto:example@example.com' style='color:blue;'>example@example.com</a> or <a href='mailto:example@example.com style='color:blue;'>example@example.com</a> <br> <b>Phone:</b> +977-1-XXXXXX Transaction Banking <br> <b>XXXXX Connect:</b> +977-1-XXXXXX <br> <b>Whatsapp/Viber:</b> 9XXXXXXXXXX</p>" +
                                "<p>Thank you for your prompt attention to this matter. We appreciate your continued trust in XXXXX XXX Bank.</p>" +
                                "<p style='color:red; font-weight:bold;'>THIS IS A SYSTEM GENERATED EMAIL. PLEASE DO NOT REPLY TO THIS EMAIL ID.</p>" +
                                "<p style='color:red; font-weight:bold;'>PLEASE DISCARD THIS MAIL IF YOU HAVE ALREADY RECEIVED THIS MAIL.</p>" +
                                "<p>Best Regards,</p>" +
                                "</body></html>", formattedDate // Use formattedDate
                            ),
                            IsBodyHtml = true, // Set to true to indicate that the body is HTML
                        };
                        // Add the recipient's email address
                        mailMessage.To.Add(recipient.Email); // Add the recipient's email here
                        
                        // Create a new attachment with the account number as the name
                        string newAttachmentName = string.Format("{0}.pdf", recipient.AccountNumber);
                        using (var stream = new FileStream(recipient.Attachment, FileMode.Open, FileAccess.Read))
                        {
                            // Create a new attachment with the new name
                            var attachment = new Attachment(stream, newAttachmentName, "application/pdf");
                            mailMessage.Attachments.Add(attachment);

                            Console.WriteLine(string.Format("Preparing to send email to {0} Attaching File {1}", recipient.Email, newAttachmentName));
                            logWriter.WriteLine(string.Format("Preparing to send email to {0} Attaching File {1} at Time {2}:", recipient.Email, newAttachmentName, DateTime.Now));

                            try
                            {
                                smtpClient.Timeout = 120000; //2 minutes
                                smtpClient.Send(mailMessage);
                                Console.WriteLine(string.Format("Email sent successfully to {0}", recipient.Email));
                                logWriter.WriteLine(string.Format("{0}: Email sent successfully to {1}", DateTime.Now, recipient.Email));
                                emailStatus = true; // Set emailStatus to true if email is sent successfully
 
                            }
                            catch (SmtpException smtpEx)
                            {
                                Console.WriteLine(string.Format("Failed to send email to {0}: {1}", recipient.Email, smtpEx.Message));
                                //Console.WriteLine(string.Format("Failed to send email to {0}: {1}", recipient.Email, smtpEx.Message));
                                logWriter.WriteLine(string.Format("{0}: Failed to send email to {1}: {2}", DateTime.Now, recipient.Email, smtpEx.Message));
                                // You can also log the status code and inner exception if available
                                //Console.WriteLine(string.Format("SMTP Status Code: {0}", smtpEx.StatusCode));
                                logWriter.WriteLine(string.Format("{0}: SMTP Status Code: {1}", DateTime.Now, smtpEx.StatusCode));

                                if (smtpEx.InnerException != null)
                                {
                                    //Console.WriteLine(string.Format("Inner Exception: {0}", smtpEx.InnerException.Message));
                                    logWriter.WriteLine(string.Format("{0}: Inner Exception: {1}", DateTime.Now, smtpEx.InnerException.Message));
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(string.Format("Failed to send email to {0}: {1}", recipient.Email, ex.Message));
                                logWriter.WriteLine(string.Format("{0}: Failed to send email to {1}: {2}", DateTime.Now, recipient.Email, ex.Message));
                            }
                        }
                        if (emailStatus == true)
                        {
                            // Attempt to delete the file after sending the email
                            try
                            {
                                // Construct the new file name
                                string movedAttachmentName = string.Format("{0};{1};{2};{3}.pdf", recipient.Email, recipient.AccountNumber, recipient.Date, recipient.CardCode);
                               // string movedAttachmentName = string.Format("{0};{1};{2}.pdf", recipient.Email, recipient.AccountNumber, recipient.Date);

                                // Construct the destination file path with the new name
                                string destinationFilePath = Path.Combine(destinationDirectory, movedAttachmentName);

                                // Move the file to the destination directory and rename it
                                File.Move(recipient.Attachment, destinationFilePath);
                                Console.WriteLine(string.Format("File {0}.pdf moved successfully.", recipient.AccountNumber));
                                logWriter.WriteLine(string.Format("{0}: File {1};{2};{3};{4}.pdf moved to {5} successfully.", DateTime.Now, recipient.Email, recipient.AccountNumber, recipient.Date, recipient.CardCode, destinationDirectory));
                            }
                            catch (IOException ioEx)
                            {
                                Console.WriteLine(string.Format("Failed to move file {0}: {1}", recipient.AccountNumber, ioEx.Message));
                                logWriter.WriteLine(string.Format("{0}: Failed to move file {1}: {2}", DateTime.Now, recipient.AccountNumber, ioEx.Message));
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine(string.Format("Attachment not found for {0} AccountNumber. Email not sent.", recipient.AccountNumber));
                        logWriter.WriteLine(string.Format("{0}: Attachment not found for {1} AccountNumber Email Address: {2}. Email not sent.", DateTime.Now, recipient.AccountNumber, recipient.Email));
                    }
                }
                // Start a task to wait for 7 seconds
                var delayTask = Task.Delay(7000);
                int countdown = 7;

                // Start a separate task for the countdown
                var countdownTask = Task.Run(async () =>
                {
                    while (countdown > 0 && !delayTask.IsCompleted)
                    {
                        Console.Write(string.Format("\rPress any key to exit or wait for {0} seconds remaining...   ", countdown)); // Overwrite the line
                        await Task.Delay(1000); // Wait for 1 second
                        countdown--;
                    }
                });

                // Wait until either a key is pressed or the delay is completed
                while (!delayTask.IsCompleted)
                {
                    if (Console.KeyAvailable)
                    {
                        Console.ReadKey(true); // Read the key without displaying it
                        break; // Exit the loop if a key is pressed
                    }
                }
                logWriter.WriteLine(string.Format("Process completed on {0}. Waiting for user input or 7 secs to exit.", DateTime.Now));

            }
        }

        // Method to validate email format
        private static bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }            
        }

        // Method to read password without echoing
        private static string ReadPassword()
        {
            string password = string.Empty;
            ConsoleKeyInfo key;

            do
            {
                key = Console.ReadKey(intercept: true); // Read key without displaying it

                // Backspace should remove a character from the password
                if (key.Key == ConsoleKey.Backspace && password.Length > 0)
                {
                    password = password.Substring(0, password.Length - 1); // Remove the last character
                    Console.Write("\b \b"); // Erase the last asterisk
                }
                // Ignore other keys
                else if (key.Key != ConsoleKey.Enter)
                {
                    password += key.KeyChar; // Add the character to the password
                    Console.Write("*"); // Display an asterisk
                }
            } while (key.Key != ConsoleKey.Enter);

            Console.WriteLine(); // Move to the next line
            return password;
        }

    }
}