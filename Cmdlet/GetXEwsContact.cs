namespace XEws.Cmdlet
{
    using System.Management.Automation;
    using CmdletAbstract;

    [Cmdlet(VerbsCommon.Get, "XEwsContact")]
    public class GetXEwsContact : XEwsContactCmdlet
    {
        protected override void ProcessRecord()
        {
            this.GetContact();
        }
    }
}
