namespace XEws.Model
{
    using System.Collections.Generic;
    using Microsoft.Exchange.WebServices.Data;

    public class EwsItem : EwsSearcher<Item>
    {
        #region Properties
        
        #endregion

        #region Fields

        private FolderId searchFolder;
        private ItemView searchItemView;

        #endregion

        #region Constructors

        public EwsItem(ExchangeService ewsSession, FolderId searchFolder, ItemView itemView, SearchFilter itemSearchFilter)
            : base(ewsSession, itemSearchFilter)
        {
            this.searchFolder = searchFolder;
            this.searchItemView = itemView;
        }

        #endregion

        #region Public Methods

        public void FindItem(FolderId rootFolder)
        {
            foreach (Item item in this.GetItem(rootFolder))
                this.OnSearchResultFound(item);
        }

        #endregion

        #region Private Methods

        private IEnumerable<Item> GetItem(FolderId searchFolder)
        {
            FindItemsResults<Item> findItemResults;

            // If no item return null, otherwise yield all results.
            do
            {
                findItemResults = ewsSession.FindItems(searchFolder, this.searchFilter, this.searchItemView);

                if (findItemResults.TotalCount == 0)
                    yield return null;
                else
                {
                    foreach (Item item in findItemResults)
                    {
                        yield return item;
                    }
                }

                // Increase offset by the value of page size
                this.searchItemView.Offset += this.searchItemView.PageSize;

            } while (findItemResults.MoreAvailable);
            
        }

        #endregion

        #region Override Methods

        #endregion
    }
}
