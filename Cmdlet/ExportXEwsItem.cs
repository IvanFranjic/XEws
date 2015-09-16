namespace XEws.Cmdlet
{
    using System.Management.Automation;
    using XEws.CmdletAbstract;

    [Cmdlet(VerbsData.Export, "XEwsItem")]
<<<<<<< HEAD
    public sealed class ExportXEwsItem : XEwsExportCmdlet
=======
    public class ExportXEwsItem : XEwsExportCmdlet
>>>>>>> 8ef1288d9a447ab7ac19a8be037d3d1bdadf5bf1
    {
        protected override void ProcessRecord()
        {
            this.ExportItem();
        }
    }
}