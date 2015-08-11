using System;
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
