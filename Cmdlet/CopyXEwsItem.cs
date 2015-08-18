namespace XEws.Cmdlet
{
    using System.Management.Automation;
    using Microsoft.Exchange.WebServices.Data;
    using XEws.CmdletAbstract;

    [Cmdlet(VerbsCommon.Copy, "XEwsItem")]
    public class CopyXEwsItem : XEwsItemMoveCmdlet
    {

        protected override void ProcessRecord()
        {
            if (this.DestinationFolder == null)
                this.DestinationFolder = XEwsFolderCmdlet.GetWellKnownFolder(WellKnownFolderName.Inbox, this.GetSessionVariable());
            try
            {
                this.MoveItem(this.Item, this.DestinationFolder, MoveOperation.Copy);
                WriteVerbose(string.Format("Copied item '{0}' to folder '{1}'.", this.Item.Subject, this.DestinationFolder.DisplayName));
            }
            catch
            {
                throw;
            }            
        }
    }
}
