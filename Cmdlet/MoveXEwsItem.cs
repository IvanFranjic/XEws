namespace XEws.Cmdlet
{
    using System.Management.Automation;
    using Microsoft.Exchange.WebServices.Data;
    using XEws.CmdletAbstract;

    [Cmdlet(VerbsCommon.Move, "XEwsItem")]
    public sealed class MoveXEwsItem : XEwsItemMoveCmdlet
    {

        protected override void ProcessRecord()
        {
            if (this.DestinationFolder == null)
                this.DestinationFolder = XEwsFolderCmdlet.GetWellKnownFolder(WellKnownFolderName.Inbox, this.GetSessionVariable());

            try
            {
                this.MoveItem(this.Item, this.DestinationFolder, MoveOperation.Move);
                WriteVerbose(string.Format("Item '{0}' moved to folder '{1}'.", this.Item.Subject, this.DestinationFolder.DisplayName));
            }
            catch
            {
                throw;
            }
        }
    }
}
