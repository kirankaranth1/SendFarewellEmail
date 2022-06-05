
using System;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SendEmail
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create the Outlook application.
                Outlook.Application oApp = new Outlook.Application();
                var dataFile = @"C:\Users\kikarant.FAREAST\Desktop\emails.txt";
                var data = File.ReadLines(dataFile);
                string work = "";
                string workExperience = string.Empty;
                string email = "";
                string firstName = "";
                var inter = string.Empty;

                foreach (var line in data)
                {
                    var trimmed = line.Trim();
                    var tmp = trimmed.Split('\t');
                    work = tmp[1].Trim('"');
                    var col1 = tmp[0].Split('<');
                    firstName = col1[0].Split(' ')[0];
                    email = col1[1].Trim('>').Trim();

                    if (work.Contains(" to ") && !work.Contains("hackathon"))
                    {
                        inter = "projects ranging from ";
                    }
                    else
                    {
                        inter = string.Empty;
                    }

                    if (work == "NA")
                    {
                        workExperience = string.Empty;
                    }
                    else
                    {
                        workExperience = $" Working with you on {inter}{work}, " +
                            $"was among those wonderful experiences and I'm very grateful for it.";
                    }

                    var subject = $"{firstName}, it was a pleasure working with you.";
                    var body = $@"Hi {firstName},<br><br>
After six wonderful years, I have decided to pursue an opportunity in an organization outside
Microsoft, with June 7th being my last working day here.<br><br>
My time here has been filled with amazing experiences that have helped me become better both as an engineer
and as a person.{workExperience} The interactions that we had is something that I will always cherish.<br><br>
I could not have asked for a better set of individuals to work with here and would like to remain in touch with you. 
Please add me on <a href=""https://www.linkedin.com/in/kirankaranth1/"">Linkedin</a> if we're not already connected there, 
and <a href=""https://linktr.ee/kikarant"">here are all my socials</a>.

<br><br>
Hope our paths cross again,<br>
Kiran<br>
+91-9740150241<br>
kirankaranth1@gmail.com";

                    SendEmail(oApp, body, subject, email);
                    Console.WriteLine($"Sending email to {email}. Subject: {subject}");
                    Console.WriteLine(body);
                    Console.WriteLine("\n\n");
                }
            }//end of try block
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        public static void SendEmail(Outlook.Application app, string body, string subject, string email = "kikarant@microsoft.com")
        {
            // Create a new mail item.
            Outlook.MailItem oMsg = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
            // Set HTMLBody. 
            //add the body of the email
            oMsg.HTMLBody = body;
            //Subject line
            oMsg.Subject = subject;
            // Add a recipient.
            Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
            // Change the recipient in the next line if necessary.
            Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(email);
            oRecip.Resolve();

            oMsg.CC = "kirankaranth1@gmail.com";
            // Send.
            oMsg.Send();
            // Clean up.
            oRecip = null;
            oRecips = null;
            oMsg = null;
            //app = null;
        }
    }
}
