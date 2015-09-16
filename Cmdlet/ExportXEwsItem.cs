namespace XEws.Cmdlet
{
    using System.Management.Automation;
    using XEws.CmdletAbstract;

    [Cmdlet(VerbsData.Export, "XEwsItem")]
    public sealed class ExportXEwsItem : XEwsExportCmdlet
    {
        protected override void ProcessRecord()
        {
            this.ExportItem();
        }
    }

}