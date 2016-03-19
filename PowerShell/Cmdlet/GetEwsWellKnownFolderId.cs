namespace XEws.PowerShell.Cmdlet
{
    using System.Management.Automation;
    using Microsoft.Exchange.WebServices.Data;

    [Cmdlet(VerbsCommon.Get, "EwsWellKnownFolderId")]
    public class GetEwsWellKnownFolderId : EwsCmdlet
    {
        #region Properties

        private WellKnownFolderName folderName = WellKnownFolderName.MsgFolderRoot;
        [Parameter(Mandatory = false, Position = 0)]
        public WellKnownFolderName FolderName
        {
            get { return this.folderName; }
            set { this.folderName = value; }
        }

        #endregion

        #region Fields

        #endregion

        #region Constructors

        #endregion

        #region Public Methods

        #endregion

        #region Private Methods

        #endregion

        #region Override Methods

        protected override void ProcessRecord()
        {
            this.WriteObject(EwsWellKnownFolder.GetWellKnownFolderId(this.folderName));            
        }

        #endregion
    }
}
