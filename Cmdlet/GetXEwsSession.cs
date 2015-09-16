namespace XEws.Cmdlet
{
    using System.Management.Automation;
    using XEws.CmdletAbstract;

    [Cmdlet(VerbsCommon.Get, "XEwsSession")]
    public sealed class GetXEwsSession : XEwsCmdlet
    {
        protected override void ProcessRecord()
        {
            WriteObject(this.GetSessionVariable());
        }
    }
}
