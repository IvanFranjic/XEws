namespace XEws.Model
{
    using System.IO;
    using Microsoft.Exchange.WebServices.Data;

    public class EwsMailTip : EwsSoapMessage
    {
        public EwsMailTipType MailTipType { get; private set; }

        public string Sender { get; private set; }

        public string Recipient { get; private set; }

        public ExchangeVersion RequestedServerVersion { get; private set; }

        public EwsMailTip(EwsMailTipType mailTipType, string sender, string recipient, ExchangeVersion requestedServerVersion)
        {
            this.MailTipType = mailTipType;
            this.Sender = sender;
            this.Recipient = recipient;
            this.RequestedServerVersion = requestedServerVersion;
        }

        // TODO: remove hardcoded path when finish with testing.
        public string ReadMailTipMessage()
        {
            string mailTipRequest = string.Empty;
            StreamReader streamReader = new StreamReader(
                @"C:\Users\ivfranji\Source\Repos\XEws\EwsMailTipRequest.xml");

            mailTipRequest = streamReader.ReadToEnd();
            mailTipRequest = 
                mailTipRequest.Replace(
                    "$requestedServerVersion",
                    this.RequestedServerVersion.ToString()).Replace(
                        "$sender",
                        this.Sender).Replace(
                            "$recipient",
                            this.Recipient).Replace(
                                "$mailTipType",
                                this.MailTipType.ToString()
                            );

            return mailTipRequest;
        }
    }
}
