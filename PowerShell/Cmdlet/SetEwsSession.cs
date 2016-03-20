namespace XEws.PowerShell.Cmdlet
{
    using System.Management.Automation;

    [Cmdlet(VerbsCommon.Set, "EwsSession")]
    public sealed class SetEwsSession : EwsSessionCmdlet
    {
        #region Properties

        private string impersonatedEmailAddress;
        [Parameter(Mandatory = false, Position = 0)]
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

        #endregion

        #region Fields

        #endregion

        #region Constructors

        #endregion

        #region Public Methods

        #endregion

        #region Private Methods

        #endregion

        #region Override Methods

        protected override void BeginProcessing()
        {
            // TODO: Implement error strings.
            if (this.EwsSession == null)
                throw new PSInvalidOperationException("No session found. Please run Import-EwsSession and try again.");
        }

        protected override void ProcessRecord()
        {
            this.SetEwsSession(this.impersonatedEmailAddress);
        }

        #endregion
    }
}
