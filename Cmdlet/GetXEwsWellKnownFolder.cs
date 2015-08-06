using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XEws.Cmdlet;
using System.Management.Automation;
using Microsoft.Exchange.WebServices.Data;
using XEws.CmdletAbstract;

namespace XEws.Cmdlet
{
    /*

        Cmdlet is returning Folder object of WellKnownFolderId. This 
        is usefull if one want get Inbox, SentItems, etc... folder
        object right away.

        Command is completed together with help.

    */

    [Cmdlet(VerbsCommon.Get, "XEwsWellKnownFolder")]
    public class GetXEwsWellKnownFolder : XEwsCmdlet
    {
        [Parameter(Mandatory = true, Position = 0)]
        public WellKnownFolderName WellKnownFolder
        {
            get;
            set;
        }

        protected override void ProcessRecord()
        {
            Folder wellKnownFolder = XEwsFolderCmdlet.GetWellKnownFolder(this.WellKnownFolder, this.GetSessionVariable());

            WriteObject(wellKnownFolder);
        }
    }
}