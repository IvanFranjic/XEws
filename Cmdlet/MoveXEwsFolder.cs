namespace XEws.Cmdlet
{
    using System;
    using Microsoft.Exchange.WebServices.Data;
    using XEws.CmdletAbstract;
    using System.Management.Automation;
    // Removing command from the module, new set of Ews command will be
    // introduced.
    //[Cmdlet(VerbsCommon.Move, "XEwsFolder")]
    public sealed class MoveXEwsFolder : XEwsFolderCmdlet
    {
        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
        public Folder Folder
        {
            get;
            set;
        }

        [Parameter(Mandatory = true, Position = 1, ValueFromPipeline = true)]
        public Folder DestinationFolder
        {
            get;
            set;
        }

        protected override void ProcessRecord()
        {
            try
            {
                this.MoveFolder(this.Folder, this.DestinationFolder);
                WriteVerbose(String.Format("Folder '{0}' moved to '{1}'.", this.Folder.DisplayName, this.DestinationFolder.DisplayName));
            }
            catch
            {
                throw;
            }            
        }

        #region Base properties overrides
        new private Folder FolderRoot
        {
            get;
            set;
        }

        new private SwitchParameter Recurse
        {
            get;
            set;
        }
        #endregion
    }
}
