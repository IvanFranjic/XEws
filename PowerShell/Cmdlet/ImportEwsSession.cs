namespace XEws.PowerShell.Cmdlet
{
    using System;
    using XEws.Model;
    using System.Net;
    using System.Security;
    using System.Diagnostics;
    using System.Management.Automation;
    using Microsoft.Exchange.WebServices.Data;
    using Microsoft.Exchange.WebServices.Autodiscover;

    [Cmdlet(VerbsData.Import, "EwsSession")]
    [CmdletBinding()]
    public sealed class ImportEwsSession : EwsSessionCmdlet
    {
        #region Properties

        private string emailAddress;
        [Parameter(Mandatory = true, Position = 0)]
        public string EmailAddress
        {
            get
            {
                return this.emailAddress;
            }
            set
            {
                this.ValidateEmailAddress(value.ToString());
                this.SetPowerShellVariable(EwsPowerShellVariable.EwsUserName, value);
                this.emailAddress = value;
            }
        }

        private SecureString password;
        [Parameter(Mandatory = true, Position = 1)]
        public SecureString Password
        {
            get
            {
                return this.password;
            }
            set
            {
                this.SetPowerShellVariable(EwsPowerShellVariable.EwsPassword, value);
                this.password = value;
            }
        }

        private string userName = null;
        [Parameter(Mandatory = false)]
        public string UserName
        {
            get
            {
                return this.userName;
            }
            set
            {
                this.userName = value;
            }
        }

        [Parameter(Mandatory = false, Position = 2, ParameterSetName = "ManualEwsEndpoint")]
        public Uri EwsEndpoint { get; set; }

        private string impersonatedEmailAddress;
        [Parameter(Mandatory = false, Position = 3)]
        public string ImpersonatedEmailAddress
        {
            get
            {
                return this.impersonatedEmailAddress;
            }
            set
            {
                this.ValidateEmailAddress(value.ToString());
                this.impersonatedEmailAddress = value;
            }
        }

        private bool traceEnabled = false;
        [Parameter(Mandatory = false, Position = 4)]
        public bool TraceEnabled
        {
            get { return this.traceEnabled; }
            set { this.traceEnabled = value; }
        }

        private TraceFlags traceFlag = TraceFlags.None;
        [Parameter(Mandatory = false, Position = 5)]
        public TraceFlags TraceFlag
        {
            get { return this.traceFlag; }
            set { this.traceFlag = value; }
        }

        private ExchangeVersion exchangeVersion = ExchangeVersion.Exchange2013_SP1;
        [Parameter(Mandatory = false)]
        public ExchangeVersion ExchangeVersion
        {
            get { return this.exchangeVersion; }
            set { this.exchangeVersion = value; }
        }

        #endregion

        #region Fields

        private Stopwatch stopWatch = new Stopwatch();

        #endregion

        #region Constructors

        #endregion

        #region Public Methods

        #endregion

        #region Private Methods

        #endregion

        #region Override Methods
        
        protected override void ProcessRecord()
        {
            stopWatch.Start();

            this.WriteWarning(
                string.Format("Instantiating ews session. Please wait...")
            );
            string userIdentity = string.Empty;
            // TODO: X-AnchorMailbox - what about impersonation. Check if impersonation is called
            // if so, anchor should be impersonated user, otherwise authenticated user.
            ExchangeService ewsSession = new ExchangeService(this.ExchangeVersion);
            this.ValidateUserName(this.EmailAddress, this.UserName, out userIdentity);
            ewsSession.Credentials = new NetworkCredential(userIdentity, this.Password);
            ewsSession.HttpHeaders.Add("X-AnchorMailbox", this.EmailAddress);
            
            if (this.traceEnabled)
            {
                ewsSession.TraceEnabled = this.traceEnabled;
                ewsSession.TraceFlags = this.TraceFlag;
                ewsSession.TraceListener = new EwsTracer();
            }

            if (this.EwsEndpoint != null)
                ewsSession.Url = this.EwsEndpoint;
            else
                ewsSession.AutodiscoverUrl(this.EmailAddress, this.AutoDiscoverRedirectValidation);

            // Are we using imperonation?
            if (!string.IsNullOrEmpty(this.ImpersonatedEmailAddress))
                ewsSession.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, this.ImpersonatedEmailAddress);

            this.SetPowerShellVariable(EwsPowerShellVariable.ExchangeService, ewsSession);

            // TODO: insert stopwatch supporting data
            this.WriteVerbose(
                string.Format("Ews session set. Endpoint used: " + this.EwsSession.Url.ToString())
            );
            // TODO: supporting data / strings class.
            this.WriteWarning("Connection established in " + this.stopWatch.Elapsed.TotalSeconds);
        }

        protected override void EndProcessing()
        {
            this.stopWatch.Stop();
        }

        #endregion
    }
}
