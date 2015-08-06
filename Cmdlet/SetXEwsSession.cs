using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using XEws.CmdletAbstract;

namespace XEws.Cmdlet
{
    [Cmdlet(VerbsCommon.Set, "XEwsSession")]
    public class SetXEwsSession : XEwsSessionCmdlet
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

        new private string Password
        {
            get;
            set;
        }

        new private Uri EwsUri
        {
            get;
            set;
        }

        #endregion
    }
}
