using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using Microsoft.Exchange.WebServices.Data;
using XEws.CmdletAbstract;

namespace XEws.Cmdlet
{
    [Cmdlet(VerbsCommon.Move, "XEwsItem")]
    public class MoveXEwsItem : XEwsItemMoveCmdlet
    {

        protected override void ProcessRecord()
        {
            if (this.DestinationFolder == null)
                this.DestinationFolder = XEwsFolderCmdlet.GetWellKnownFolder(WellKnownFolderName.Inbox, this.GetSessionVariable());

            try
            {
                this.MoveItem(this.Item, this.DestinationFolder, MoveOperation.Move);
                WriteVerbose(string.Format("Moved item '{0}' to folder '{1}'.", this.Item.Subject, this.DestinationFolder.DisplayName));
            }
            catch
            {
                throw;
            }
        }
    }
}
