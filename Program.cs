using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using MimeKit;
using MailKit;
using MailKit.Net.Smtp;

namespace TestEmailer
{
    class Program
    {
        protected static string HostIP = ConfigurationManager.AppSettings["SMTPip"].ToString();
        protected static string Username = ConfigurationManager.AppSettings["SMTPusername"].ToString();
        protected static string Pass = ConfigurationManager.AppSettings["SMTPpassword"].ToString();
        protected static int HostPort = Int32.Parse(ConfigurationManager.AppSettings["SMTPport"]);
        protected static string FromEmail = ConfigurationManager.AppSettings["EmailFrom"].ToString();
        protected static string Recipient = ConfigurationManager.AppSettings["EmailNotificationTo"].ToString();
        protected static string SubjectStr = "Mailkit Test";

        static void Main()
        {
            // Test using MailKit and MimeKit
            MimeMessage mimeMail = CreateNewMimeMessage();
            SendMimeMessage(mimeMail);

            // Test using Mail Message and depreciated SmtpClient
            //MailMessage mail = CreateNewMailMessage();
            //SendMailMessage(mail);

            Console.WriteLine("Mail Sent!");
            Console.ReadLine();
        }

        public static MailMessage CreateNewMailMessage()
        {
            MailMessage mm = new MailMessage();
            mm.From = new MailAddress(FromEmail);
            mm.Subject = SubjectStr;
            mm.Body = GetMailMessage();
            mm.IsBodyHtml = true;
            mm.To.Add(Recipient);
            return mm;
        }

        public static string GetMailMessage()
        {
            string testMessage = "Someday we'll find it the Rainbow Connection The lovers, the dreamers, and me";
            string[] testWords = testMessage.Split(' ');
            StringBuilder HTMLBody = new StringBuilder();
            HTMLBody.Append("<HTML>");
            HTMLBody.Append("<HEAD>");
            HTMLBody.Append("<TITLE>" + SubjectStr +"</TITLE>");
            HTMLBody.Append("</HEAD>");
            HTMLBody.Append("<BODY>");
            HTMLBody.Append("<H1>Why are there so many songs about rainbows</H1>");
            HTMLBody.Append("<H2>And what's on the other side?</H2>");
            HTMLBody.Append("<P>Rainbows are visions, they're only illusions and rainbows have nothing to hide.  So we've been told and some chose to believe it, but I know they're wrong wait and see.</P>");
            HTMLBody.Append("<TABLE width=600>");
            HTMLBody.Append("<TR>");

            HTMLBody.Append("<TD width='200px'>");
            HTMLBody.Append("Column 1");
            HTMLBody.Append("</TD>");

            HTMLBody.Append("<TD width='200px'>");
            HTMLBody.Append("Column 2");
            HTMLBody.Append("</TD>");

            HTMLBody.Append("<TD width='200px'>");
            HTMLBody.Append("Column 3");
            HTMLBody.Append("</TD>");

            HTMLBody.Append("</TR>");

            HTMLBody.Append("<TR>");

            for (int i = 0; i < testWords.Length; i++)
            {
                HTMLBody.Append("<TD>");
                HTMLBody.Append(testWords[i]);
                HTMLBody.Append("</TD>");
                
                if ((i + 1)%3 == 0)
                {
                    HTMLBody.Append("</TR>");
                    if (i != testWords.Length - 1)
                    {
                        HTMLBody.Append("<TR>");
                    }
                }

            }

            HTMLBody.Append("</TABLE>");
            HTMLBody.Append("</BODY>");
            HTMLBody.Append("</HTML>");
            return HTMLBody.ToString();
        }

        public static void SendMailMessage (MailMessage mail)
        {
            var smtp = new System.Net.Mail.SmtpClient();
            smtp.Host = HostIP;

            smtp.EnableSsl = true;
            NetworkCredential netCred = new NetworkCredential();
            netCred.UserName = Username;
            netCred.Password = Pass;
            smtp.UseDefaultCredentials = true;
            smtp.Credentials = netCred;
            smtp.Port = HostPort;
            smtp.Send(mail);
        }

        public static MimeMessage CreateNewMimeMessage()
        {
            MimeMessage mm = new MimeMessage();
            mm.From.Add(MailboxAddress.Parse(FromEmail));
            mm.Subject = SubjectStr;
            mm.To.Add(MailboxAddress.Parse(Recipient));
            BodyBuilder bb = new BodyBuilder();
            bb.HtmlBody = GetMailMessage();
            mm.Body = bb.ToMessageBody();
            return mm;
        }

        public static void SendMimeMessage (MimeMessage mail)
        {
            var client = new MailKit.Net.Smtp.SmtpClient();
            client.Connect(HostIP, HostPort, false);
            client.Authenticate(Username, Pass);
            client.Send(mail);
            client.Disconnect(true);
        }
    }
}
