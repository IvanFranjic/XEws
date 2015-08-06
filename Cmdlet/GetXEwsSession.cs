using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using XEws.CmdletAbstract;

namespace XEws.Cmdlet
{
    [Cmdlet(VerbsCommon.Get, "XEwsSession")]
    public class GetXEwsSession : XEwsCmdlet
    {
        protected override void ProcessRecord()
        {
            WriteObject(this.GetSessionVariable(), true);
        }
    }
}
