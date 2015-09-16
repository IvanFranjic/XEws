namespace XEws.Cmdlet
{
    using System;
    using System.Management.Automation;
    using XEws.CmdletAbstract;
    using System.Security;

    [Cmdlet(VerbsCommon.Set, "XEwsSession")]
    public sealed class SetXEwsSession : XEwsSessionCmdlet
    {
        protected override void ProcessRecord()
        {
            if (this.TraceEnabled)
            {
                this.ValidateTracePath(this.TraceOutputFolder);
            }

            this.SetSessionVariable(this.ImpersonationEmail, this.TraceEnabled, this.TraceOutputFolder, this.TraceFlags);
        }

        #region Override parameters from base
        
        new private string UserName
        {
            get;
            set;
        }

        new private SecureString Password
        {
            get;
            set;
        }

        new private Uri EwsUri
        {
            get;
            set;
        }

        new private string AutodiscoverEmail
        {
            get;
            set;
        }

        #endregion
    }
}
