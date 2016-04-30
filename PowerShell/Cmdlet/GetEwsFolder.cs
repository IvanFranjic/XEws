namespace XEws.PowerShell.Cmdlet
{
    using XEws.Model;
    using System.Management.Automation;
    using Microsoft.Exchange.WebServices.Data;

    [Cmdlet(VerbsCommon.Get, "EwsFolder")]
    [CmdletBinding()]
    public class GetEwsFolder : EwsCmdlet
    {
        #region Properties

        private FolderId rootFolder = EwsWellKnownFolder.GetWellKnownFolderId(WellKnownFolderName.MsgFolderRoot);
        [Parameter(Position = 0)]
        public FolderId RootFolder
        {
            get
            {
                return this.rootFolder;
            }
            set
            {
                this.rootFolder = value;
            }
        }

        private string folderName = "Inbox";
        [Parameter(Position = 1, ParameterSetName = "FolderName")]
        [ValidateNotNullOrEmpty()]
        public string FolderName
        {
            get
            {
                return this.folderName;
            }
            set
            {
                this.folderName = value;
            }
        }

        private bool recurse = false;
        //[Parameter(Mandatory = false, ParameterSetName = "RecurseSearch")]
        //public bool Recurse
        //{
        //    get
        //    {
        //        return this.recurse;
        //    }
        //    set
        //    {
        //        this.recurse = value;
        //    }
        //}


        #endregion

        #region Fields

        // TODO: implement searchfilter for advanced
        // folder search.
        /// <summary>
        /// Not implemented at the moment.
        /// </summary>
        private SearchFilter searchFilter;

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
            EwsFolder folderSearcher = new EwsFolder(this.EwsSession, null);
            folderSearcher.SearchResultFound += OnSearchResultFound;

            switch (this.ParameterSetName)
            {
                case "FolderName":
                    folderSearcher.FindFolderByName(this.RootFolder, this.FolderName);
                    break;

                default:
                    folderSearcher.FindFolder(this.RootFolder, this.recurse);
                    break;
            }
        }

        /// <summary>
        /// This will be subscribed to SearchResultFound event.
        /// </summary>
        /// <param name="sender">Sender of the event.</param>
        /// <param name="e">EwsSearchResultEntryEventArgs<Folder></Item></param>
        private void OnSearchResultFound(object sender, EwsSearchResultEntryEventArgs<Folder> e)
        {
            // If we are subscribed to event, output object to the PS console.
            this.WriteObject(e.SearchResultEntry);
        }

        #endregion
    }
}