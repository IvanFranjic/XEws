namespace XEws.Cmdlet
{
    using System;
    using Microsoft.Exchange.WebServices.Data;
    using XEws.CmdletAbstract;
    using System.Management.Automation;
    // Removing command from the module, new set of Ews command will be
    // introduced.
    //[Cmdlet(VerbsCommon.Remove, "XEwsFolder")]
    public sealed class RemoveXEwsFolder : XEwsFolderCmdlet
    {
        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
        public Folder Folder
        {
            get;
            set;
        }

        [Parameter(Mandatory = true, Position = 1)]
        public DeleteMode DeleteMode
        {
            get;
            set;
        }

        [Parameter(Mandatory = false, Position = 2)]
        public SwitchParameter Force
        {
            get;
            set;
        }

        protected override void ProcessRecord()
        {
            try
            {
                this.RemoveFolder(this.Folder, this.DeleteMode, this.Force);
                WriteVerbose(String.Format("Successfuly removed folder {0}.", this.Folder.DisplayName));
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
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
