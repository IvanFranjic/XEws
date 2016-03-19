namespace XEws.Model
{
    using System.Collections.Generic;
    using Microsoft.Exchange.WebServices.Data;
    
    public class EwsFolder : EwsSearcher<Folder>
    {
        #region Properties

        #endregion

        #region Fields

        #endregion

        #region Constructors
        // TODO: implement Searchfilter
        /// <summary>
        /// Instantiate EwsFolder class.
        /// </summary>
        /// <param name="ewsSession">Exchange service used to bind to the EWS endpoint.</param>
        /// <param name="folderView">Folder view.</param>
        public EwsFolder(ExchangeService ewsSession, SearchFilter folderSearchFilter)
            : base (ewsSession, folderSearchFilter)
        {
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Find all folder under specified root. If folder is found it will
        /// be passed to SearchResultFound event. If there is no
        /// subscription to event nothing will be returned.
        /// </summary>
        /// <param name="parentFolder">Root folder from where search should start.</param>
        public void FindFolder(FolderId rootFolder, bool searchRecurse)
        {
            foreach (Folder folder in this.GetFolder(rootFolder, new FolderView(10)))
            {
                this.OnSearchResultFound(folder);

                // We should check if both recurse is called and 
                // if current folder actually have child folders.
                if (searchRecurse && folder.ChildFolderCount > 0)
                    this.FindFolder(folder.Id, searchRecurse);
            }
        }

        /// <summary>
        /// Finds folder specified by name. If folder is found it will
        /// be passed to SearchResultFound event. If there is no
        /// subscription to event nothing will be returned.
        /// </summary>
        /// <param name="parentFolder">Root folder from where search should start.</param>
        /// <param name="folderName">Name of the folder to find.</param>
        public void FindFolderByName(FolderId rootFolder, string folderName)
        {
            foreach (Folder folder in this.GetFolder(rootFolder, new FolderView(10)))
            {
                if (folder.DisplayName == folderName)
                {
                    this.OnSearchResultFound(folder);
                    break;
                }
            }
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Gets all folder from specific root specified.
        /// </summary>
        /// <param name="parentFolder">Root folder where search should start.</param>
        /// <returns></returns>
        private IEnumerable<Folder> GetFolder(FolderId parentFolder, FolderView folderView)
        {
            FindFoldersResults findFolderResults;

            // If no folder found return null, otherwise yield all folders found.
            do
            {
                findFolderResults = this.ewsSession.FindFolders(parentFolder, folderView);
                if (findFolderResults.TotalCount == 0)
                    yield return null;
                else
                { 
                    
                    foreach (Folder folder in findFolderResults)
                    {
                        yield return folder;
                    }
                }

                // Increase offset by the value of page size
                folderView.Offset += folderView.PageSize;

            } while (findFolderResults.MoreAvailable);

        }

        #endregion

        #region Override Methods

        #endregion
    }
}
