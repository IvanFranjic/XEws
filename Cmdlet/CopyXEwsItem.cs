﻿namespace XEws.Cmdlet
{
    using System.Management.Automation;
    using Microsoft.Exchange.WebServices.Data;
    using XEws.CmdletAbstract;
    // Removing command from the module, new set of Ews command will be
    // introduced.
    //[Cmdlet(VerbsCommon.Copy, "XEwsItem")]
    public sealed class CopyXEwsItem : XEwsItemMoveCmdlet
    {

        protected override void ProcessRecord()
        {
            if (this.DestinationFolder == null)
                this.DestinationFolder = XEwsFolderCmdlet.GetWellKnownFolder(WellKnownFolderName.Inbox, this.GetSessionVariable());
            try
            {
                this.MoveItem(this.Item, this.DestinationFolder, MoveOperation.Copy);
                WriteVerbose(string.Format("Item '{0}' copied to folder '{1}'.", this.Item.Subject, this.DestinationFolder.DisplayName));
            }
            catch
            {
                throw;
            }            
        }
    }
}
