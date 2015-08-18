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
    [Cmdlet(VerbsCommon.Get, "XEwsFolderByName")]
    public class GetXEwsFolderByName : XEwsFolderCmdlet
    { 
        [Parameter(Mandatory = true, Position = 1)]
        public string FolderName
        {
            get;
            set;
        }

        protected override void ProcessRecord()
        {
            
            /*            
                If search root is not specified, default to Inbox.
            */

            if (this.FolderRoot == null)
                this.FolderRoot = XEwsFolderCmdlet.GetWellKnownFolder(WellKnownFolderName.Inbox, this.GetSessionVariable());

            Folder folder = null;

            if (this.Recurse)
            {
                WriteWarning("Use 'Recurse' switch with a caution. If more than one folder with the same name is found, last one will be returned which can lead to undesired behaviour.");

                List<Folder> folders = new List<Folder>();
                GetRecurseFolder(this.FolderRoot, ref folders);
                
                foreach (Folder recurseFolder in folders)
                {
                    if (recurseFolder.DisplayName == this.FolderName)
                    {
                        folder = recurseFolder;
                        break;
                    }
                }
            }

            /*
                Find folder under exactly specified root. Dont do recurse.
            */

            else
            {
                folder = this.GetFolder(this.FolderRoot, this.FolderName);                
            }

            if (folder == null)
                WriteVerbose(String.Format("No folders found matching name '{0}' under '{1}'.", this.FolderName, this.FolderRoot.DisplayName.ToString()));

            WriteObject(folder, true);
        }

    }
}
