namespace XEws.Cmdlet
{
    using System.Management.Automation;
    using CmdletAbstract;
    // Removing command from the module, new set of Ews command will be
    // introduced.
    //[Cmdlet(VerbsCommon.Get, "XEwsContact")]
    public class GetXEwsContact : XEwsContactCmdlet
    {
        protected override void ProcessRecord()
        {
            this.GetContact();
        }
    }
}
