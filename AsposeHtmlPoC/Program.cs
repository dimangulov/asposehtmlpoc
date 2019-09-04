using System;
using System.Diagnostics;
using System.IO;
using Aspose.Email;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Tools.Merging;
using Aspose.Email.Tools.Merging.DataReader;
using Aspose.Html;
using Aspose.Html.Converters;
using Aspose.Html.Loading;

namespace AsposeHtmlPoC
{
    class Program
    {
        static void Main(string[] args)
        {
            // Instantiate the License class
            var license1 = new Aspose.Email.License();
            var license2 = new Aspose.Html.License();

            // license should be stored as an embedded resource
            license1.SetLicense("Aspose.Total.lic");
            license2.SetLicense("Aspose.Total.lic");

            RunEmail(RunHtml());
        }

        private static string RunHtml()
        {
            var templateFile = Path.Combine(Environment.CurrentDirectory, "template.html");
            var dataFile = Path.Combine(Environment.CurrentDirectory, "data.json");
            var templateHtml = new HTMLDocument(templateFile);
            //XML data for merging 
            var data = new TemplateData(dataFile);
            //Output file path 
            string templateOutput = @"c:\temp\messages\Merged_Output"+DateTime.Now.Ticks+".html";
            //Merge HTML tempate with XML data
            Converter.ConvertTemplate(templateHtml, data, new TemplateLoadOptions(), templateOutput);

            return templateOutput;
        }

        public static void RunEmail(string htmlFilePath)
        {
            // Create a new MailMessage instance
            MailMessage msg = new MailMessage();

            // Add subject and from address
            msg.Subject = "Hello, #FirstName#";
            msg.From = "dimangulov@gmail.com";

            // Add email address to send email also Add mesage field to HTML body
            msg.To.Add("dimangulov@gmail.com");
            msg.HtmlBody = File.ReadAllText(htmlFilePath);

            try
            {
                // Create messages from the message and datasource.

                // Create an instance of SmtpClient and specify server, port, username and password
                var client = new SmtpClient("smtp.gmail.com", 587, "{your_account}@gmail.com", "{application_password}")
                {
                    SecurityOptions = SecurityOptions.Auto,
                    UseAuthentication = true,
                    
                };

                // Send messages in bulk
                client.Send(msg);
            }
            catch (MailException ex)
            {
                Debug.WriteLine(ex.ToString());
            }
            catch (SmtpException ex)
            {
                Debug.WriteLine(ex.ToString());
            }

            Console.WriteLine(Environment.NewLine + "Message sent after performing mail merge.");
        }
    }
}
