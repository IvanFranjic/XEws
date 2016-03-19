namespace XEws.Model
{
    using System;
    using Microsoft.Exchange.WebServices.Data;

    public abstract class EwsSearcher<Type>
    {
        #region Properties

        public FolderId RootFolderId { get; private set; }

        #endregion

        #region Fields

        protected ExchangeService ewsSession;
        protected SearchFilter searchFilter;

        #endregion

        #region Constructors

        /// <summary>
        /// Instantiate EwsSearcher class.
        /// </summary>
        /// <param name="ewsService">Instance of ExchangeService.</param>
        /// <param name="rootFolderId">Root folder from where search starts.</param>
        /// <param name="searchFilter">Search filter.</param>
        public EwsSearcher(ExchangeService ewsService, SearchFilter searchFilter)
        {
            this.ewsSession = ewsService;
            this.searchFilter = searchFilter;
        }

        #endregion

        #region Public Methods

        #endregion

        #region Private Methods

        /// <summary>
        /// Method will instantiate SearchResultFound if there is
        /// subscriptions to the event.
        /// </summary>
        /// <param name="searchResult">ItemType to contained by SearchResultFound.</param>
        protected virtual void OnSearchResultFound(Type searchResult)
        {
            if (this.SearchResultFound != null)
                this.SearchResultFound(this, new EwsSearchResultEntryEventArgs<Type>(searchResult));
        }

        #endregion

        #region Override Methods

        #endregion

        #region Events and Delegates

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T">Type of event handler.</typeparam>
        /// <param name="sender">Object that broadcast event.</param>
        /// <param name="e">EwsSearchResultEntryEventArg</param>
        public delegate void SearchResultEventHandler<ItemType>(object sender, EwsSearchResultEntryEventArgs<ItemType> e);

        /// <summary>
        /// Event that will fire every time search result is found.
        /// </summary>
        public event SearchResultEventHandler<Type> SearchResultFound;

        #endregion

    }
}
