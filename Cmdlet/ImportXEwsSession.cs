﻿namespace XEws.Cmdlet
{
    using System.Management.Automation;
    using XEws.CmdletAbstract;
    using System;

    // Removing command from the module, new set of Ews command will be
    // introduced.
    //[Cmdlet("Import", "XEwsSession")]
    public sealed class ImportXEwsSession : XEwsSessionCmdlet
    {

        protected override void ProcessRecord()
        {
            this.SetSessionVariable(this.UserName, this.Password, this.EwsUri, this.AutodiscoverEmail, this.ImpersonationEmail, this.TraceEnabled, this.TraceOutputFolder, this.TraceFlags, this.exchangeVersion);

            WriteVerbose(String.Format("Using following Ews endpoint: {0}", this.GetSessionVariable().Url.ToString()));
        }
    }
}
