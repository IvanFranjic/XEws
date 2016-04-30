namespace XEws.Model
{
    using System.IO;
    using System.Xml;
    using System.Net;
    using System.Text;
    using Microsoft.Exchange.WebServices.Data;

    public class EwsMailTip : EwsSoapMessage
    {
        public EwsMailTipType MailTipType { get; private set; }

        public string Sender { get; private set; }

        public string Recipient { get; private set; }

        public string MailTipResponse { get; private set; }

        public ExchangeVersion RequestedServerVersion { get; private set; }

        public EwsMailTip(EwsMailTipType mailTipType, string sender, string recipient, ExchangeVersion requestedServerVersion)
        {
            this.MailTipType = mailTipType;
            this.Sender = sender;
            this.Recipient = recipient;
            this.RequestedServerVersion = requestedServerVersion;
        }

        
        public XmlDocument GetMailTip(ICredentials credentials, string url, bool traceEnabled)
        {
            EwsTracer tracer = new EwsTracer(traceEnabled);

            string soapMessage = this.GetSoapMessage();

            tracer.Trace("MailTipRequest", soapMessage);

            WebRequest webRequest = WebRequest.Create(url);
            webRequest.Headers.Set(HttpRequestHeader.Pragma, "no-cache");
            webRequest.Headers.Set(HttpRequestHeader.Translate, "f");
            webRequest.Headers.Add("Depth", "0");
            webRequest.ContentType = "text/xml";
            webRequest.ContentLength = soapMessage.Length;
            webRequest.Timeout = 60000;
            webRequest.Method = "POST";
            webRequest.Credentials = credentials;
            byte[] bytesQuery = Encoding.ASCII.GetBytes(soapMessage);
            webRequest.ContentLength = bytesQuery.Length;
            Stream requestStream = webRequest.GetRequestStream();
            requestStream.Write(bytesQuery, 0, bytesQuery.Length);
            requestStream.Close();

            WebResponse webResponse = webRequest.GetResponse();
            Stream webResponseStream = webResponse.GetResponseStream();

            StreamReader streamReader = new StreamReader(webResponseStream);
            XmlDocument soapResponse = new XmlDocument();
            soapResponse.LoadXml(streamReader.ReadToEnd());

            string soapResponseString = this.GetXmlText(soapResponse);

            tracer.Trace("MailTipResponse", soapResponseString);

            return soapResponse;
        }

        /// <summary>
        /// Gets string representation of xml document.
        /// </summary>
        /// <param name="xmlDocument"></param>
        /// <returns></returns>
        private string GetXmlText(XmlDocument xmlDocument)
        {
            return xmlDocument.OuterXml;

            //using (StringWriter stringWriter = new StringWriter())
            //{
            //    using (var xmlWriter = XmlWriter.Create(stringWriter))
            //    {
            //        xmlDocument.WriteTo(xmlWriter);
            //        xmlWriter.Flush();

            //        return stringWriter.GetStringBuilder().ToString();
            //    }
            //}
        }


        /// <summary>
        /// Creates MailTip soap message.
        /// </summary>
        /// <returns>string</returns>
        private string GetSoapMessage()
        {
            StringBuilder stringBuilter = new StringBuilder();

            // TODO: Remove spaces from request in xml tags since they are causing issue.
            #region Old soapRequest
            /*
            stringBuilter.AppendLine("<?xml version=\"1.0\" encoding=\"utf - 8\"?>");
            stringBuilter.AppendLine("< soap:Envelope xmlns: soap = \"http://schemas.xmlsoap.org/soap/envelope/\" xmlns: xsi = \"http://www.w3.org/2001/XMLSchema-instance\" xmlns: xsd = \"http://www.w3.org/2001/XMLSchema\" >");
            stringBuilter.AppendLine("<soap:Header>");
            stringBuilter.AppendLine("<RequestServerVersion Version=\"$requestedServerVersion\" xmlns=\"http://schemas.microsoft.com/exchange/services/2006/types\" />");
            stringBuilter.AppendLine("</soap:Header>");
            stringBuilter.AppendLine("<soap:Body>");
            stringBuilter.AppendLine("<GetMailTips xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\">");
            stringBuilter.AppendLine("<SendingAs>");
            stringBuilter.AppendLine("<EmailAddress xmlns=\"http://schemas.microsoft.com/exchange/services/2006/types\">$sender</EmailAddress>");
            stringBuilter.AppendLine("</SendingAs>");
            stringBuilter.AppendLine("<Recipients>");
            stringBuilter.AppendLine("<Mailbox xmlns=\"http://schemas.microsoft.com/exchange/services/2006/types\">");
            stringBuilter.AppendLine("<EmailAddress>$recipient</EmailAddress>");
            stringBuilter.AppendLine("</Mailbox>");
            stringBuilter.AppendLine("</Recipients>");
            stringBuilter.AppendLine("<MailTipsRequested>$mailTipType</MailTipsRequested>");
            stringBuilter.AppendLine("</GetMailTips>");
            stringBuilter.AppendLine("</soap:Body>");
            stringBuilter.AppendLine("</soap:Envelope>");

            */
            #endregion

            stringBuilter.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            stringBuilter.AppendLine("<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">");
            stringBuilter.AppendLine("<soap:Header><RequestServerVersion Version=\"$requestedServerVersion\" xmlns=\"http://schemas.microsoft.com/exchange/services/2006/types\" />");
            stringBuilter.AppendLine("</soap:Header>");
            stringBuilter.AppendLine("<soap:Body>");
            stringBuilter.AppendLine("<GetMailTips xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\">");
            stringBuilter.AppendLine("<SendingAs>");
            stringBuilter.AppendLine("<EmailAddress xmlns=\"http://schemas.microsoft.com/exchange/services/2006/types\">$sender</EmailAddress>");
            stringBuilter.AppendLine("</SendingAs>");
            stringBuilter.AppendLine("<Recipients><Mailbox xmlns=\"http://schemas.microsoft.com/exchange/services/2006/types\"><EmailAddress>$recipient</EmailAddress></Mailbox></Recipients><MailTipsRequested>$mailTipType</MailTipsRequested></GetMailTips></soap:Body></soap:Envelope>");

            stringBuilter.Replace(
                "$requestedServerVersion", this.RequestedServerVersion.ToString()).Replace(
                    "$sender", this.Sender).Replace(
                        "$recipient", this.Recipient).Replace(
                            "$mailTipType", this.MailTipType.ToString());

            return stringBuilter.ToString();

        }
        /*

            <?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
            <soap:Header><RequestServerVersion Version="$requestedServerVersion" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" />
            </soap:Header>
            <soap:Body>
            <GetMailTips xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
            <SendingAs>
            <EmailAddress xmlns="http://schemas.microsoft.com/exchange/services/2006/types">$sender</EmailAddress>
            </SendingAs>
            <Recipients><Mailbox xmlns="http://schemas.microsoft.com/exchange/services/2006/types"><EmailAddress>$recipient</EmailAddress></Mailbox></Recipients><MailTipsRequested>$mailTipType</MailTipsRequested></GetMailTips></soap:Body></soap:Envelope>


        */

    }
}
