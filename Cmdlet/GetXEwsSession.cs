namespace XEws.Cmdlet
{
    using System.Management.Automation;
    using XEws.CmdletAbstract;

    // Removing command from the module, new set of Ews command will be
    // introduced.
    //[Cmdlet(VerbsCommon.Get, "XEwsSession")]
    public sealed class GetXEwsSession : XEwsCmdlet
    {
        protected override void ProcessRecord()
        {
            WriteObject(this.GetSessionVariable());
        }
    }
}
