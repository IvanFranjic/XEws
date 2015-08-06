using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using XEws.CmdletAbstract;

namespace XEws.Cmdlet
{
    [Cmdlet("Import", "XEwsSession")]
    public class ImportXEwsSession : XEwsSessionCmdlet
    {
        protected override void ProcessRecord()
        {
            this.SetSessionVariable(this.UserName, this.Password, this.EwsUri, this.ImpersonationEmail, this.TraceEnabled, this.TraceOutputFolder, this.TraceFlags);
        }
    }
}
