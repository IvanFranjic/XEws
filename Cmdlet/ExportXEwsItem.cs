namespace XEws.Cmdlet
{
    using System.Management.Automation;
    using XEws.CmdletAbstract;

    // Removing command from the module, new set of Ews command will be
    // introduced.
    //[Cmdlet(VerbsData.Export, "XEwsItem")]
    public sealed class ExportXEwsItem : XEwsExportCmdlet
    {
        protected override void ProcessRecord()
        {
            this.ExportItem();
        }
    }
}