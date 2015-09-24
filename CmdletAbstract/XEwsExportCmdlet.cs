namespace XEws.CmdletAbstract
{
    using System;
    using System.Management.Automation;
    using Microsoft.Exchange.WebServices.Data;
    using System.IO;

    public abstract class XEwsExportCmdlet : XEwsCmdlet
    {
        [Parameter(Position = 0, ValueFromPipeline = true, Mandatory = true)]
        [ValidateNotNullOrEmpty]
        public Item ItemToExport
        {
            get;
            set;
        }

        private string fileName = null;
        [Parameter(Position = 1)]
        public string FileName
        {
            get
            {
                return fileName;
            }
            set
            {
                fileName = value;
            }
        }


        #region Methods

        internal void ExportItem()
        {
            if (this.ItemToExport.GetType().Name.ToString() != "EmailMessage")
                throw new InvalidOperationException("Please provide item typeof EmailMessage.");

            ExchangeService ewsSession = this.GetSessionVariable();
            PropertySet propertySet = new PropertySet(EmailMessageSchema.MimeContent);
            propertySet.Add(EmailMessageSchema.ConversationTopic);

            EmailMessage emailMessage = EmailMessage.Bind(ewsSession, this.ItemToExport.Id, propertySet);

            Random rNumber = new Random();
            string directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "EmailExport");

            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            if (string.IsNullOrEmpty(this.fileName))
            {
                int subjectLenght = emailMessage.ConversationTopic.Length > 10 ? 10 : emailMessage.ConversationTopic.Length;
                fileName = String.Format( "Email-{0}-{1}" , emailMessage.ConversationTopic.Substring( 0 , subjectLenght ) , rNumber.Next( 1 , 10 ) );
            }

            fileName = Path.Combine(directory, string.Format("{0}.eml", fileName));

            using (FileStream fileStream = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                fileStream.Write(emailMessage.MimeContent.Content, 0, emailMessage.MimeContent.Content.Length);
            }
        }

        #endregion
    }

}