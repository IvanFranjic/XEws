namespace XEws.PowerShell.Cmdlet
{
    using System.IO;
    using XEws.Model;
    using System.Net;
    using System.Xml;
    using System.Text;
    using System.Security;
    using System.Management.Automation;

    [Cmdlet(VerbsCommon.Get, "EwsMailTip")]
    public class GetEwsMailTip : EwsMailTipCmdlet
    {
        [Parameter(Mandatory = true)]
        public string Recipient { get; set; }

        [Parameter(Mandatory = true)]
        public EwsMailTipType MailTipType
        {
            get;
            set;
        }
        

        protected override void ProcessRecord()
        {
            EwsMailTip ewsMailTip = new EwsMailTip(
                this.MailTipType, 
                (string)this.GetPowerShellVariable(EwsPowerShellVariable.EwsUserName).Value, 
                this.Recipient, 
                this.EwsSession.RequestedServerVersion);

            NetworkCredential credentials = new NetworkCredential(
                (string)this.GetPowerShellVariable(EwsPowerShellVariable.EwsUserName).Value, 
                (SecureString)this.GetPowerShellVariable(EwsPowerShellVariable.EwsPassword).Value);
            
            XmlDocument mailTip = ewsMailTip.GetMailTip(credentials, this.EwsSession.Url.ToString(), this.EwsSession.TraceEnabled);

            this.WriteObject(mailTip, true);
        }
    }
}
