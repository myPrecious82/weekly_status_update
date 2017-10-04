using System;
using System.Configuration;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Net.Mail;

namespace weekly_status_update
{
    public class Program
    {
        public static void Main()
        {
            var appSettings = ConfigurationManager.AppSettings;

            var path = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName), @"..\..\"));
            var templatePath = $"{path}{appSettings["TemplateFileName"]}";

            object outputPath = $"{path}{appSettings["WordOutputFileName"]}";
            var newPath = outputPath.ToString().Replace(".docx", $" - {appSettings["ConsultantName"]}.docx");
            object oMissing = System.Reflection.Missing.Value;

            System.Globalization.CultureInfo ci = System.Threading.Thread.CurrentThread.CurrentCulture;
            DayOfWeek fdow = ci.DateTimeFormat.FirstDayOfWeek;
            DayOfWeek today = DateTime.Now.DayOfWeek;
            DateTime sow = DateTime.Now.AddDays(-(today - fdow)).Date;
            DateTime eow = sow.AddDays(6);
            DateTime friday = eow.AddDays(-1);
            var emailSmtpHost = appSettings["EmailSmtpHost"];
            var emailSendFrom = appSettings["EmailSendFrom"];
            var emailSendFromDisplay = appSettings["EmailSendFromDisplay"];
            string outputFileName;

            var byteArray = File.ReadAllBytes(templatePath);
            using (MemoryStream stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, (int)byteArray.Length);
                using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(stream, true))
                {
                    var document = wordprocessingDocument.MainDocumentPart.Document;

                    foreach (var text in document.Descendants<Text>())
                        switch (text.Text)
                        {
                            case "{Todays Date}":
                                text.Text = text.Text.Replace("{Todays Date}", friday.ToString("MMMM dd, yyyy"));
                                break;
                            case "{Reporting Period}":
                                text.Text = text.Text.Replace("{Reporting Period}", $"{sow:MM/dd/yyyy} - {eow:MM/dd/yyyy}");
                                break;
                            case "{Next Report Date}":
                                text.Text = text.Text.Replace("{Next Report Date}", friday.AddDays(7).ToString("MMMM dd, yyyy"));
                                break;
                            case "{Consultant Name}":
                                text.Text = text.Text.Replace("{Consultant Name}", appSettings["ConsultantName"]);
                                break;
                        }

                    outputFileName = newPath.ToString().Replace(".docx", $" {sow:MM.dd.yyyy}.docx");
                    wordprocessingDocument.Close();
                    wordprocessingDocument.Dispose();
                }

                using (FileStream fs = new FileStream(outputFileName, FileMode.OpenOrCreate))
                {
                    stream.WriteTo(fs);
                    fs.Close();
                    fs.Dispose();
                }
                stream.Close();
                stream.Dispose();
            }
            
            var emailSendTo = appSettings["EmailSendTo"];
#if DEBUG
                emailSendTo = appSettings["DebugEmailSendTo"];
#endif
            var emailSubject = appSettings["EmailSubject"];
            var client = new SmtpClient(emailSmtpHost);

            using (MailMessage message = new MailMessage())
            {
                message.IsBodyHtml = true;
                message.BodyEncoding = System.Text.Encoding.UTF8;
                message.Subject = emailSubject;
                message.SubjectEncoding = System.Text.Encoding.UTF8;
                message.Bcc.Add(new MailAddress(emailSendFrom));
                message.From = new MailAddress(emailSendFrom, emailSendFromDisplay, System.Text.Encoding.UTF8);
                foreach (var x in emailSendTo.Split(','))
                {
                    message.To.Add(new MailAddress(x));
                }
            message.Attachments.Add(new Attachment(outputFileName.ToString()));
            message.Body = $"{message.Body}<span style='font-size:11pt;font-family:Calibri'>{appSettings["EmailBody"]}</span>";

            client.Send(message);

            message.Dispose();
        }
    }
}
}
