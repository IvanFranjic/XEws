namespace XEws.Model
{
    using System;

    /// <summary>
    /// Type generic EventArgs class to all search results
    /// which could be Folder, Item, Contact...
    /// </summary>
    /// <typeparam name="T">Type of search result - Folder, Item...</typeparam>
    public sealed class EwsSearchResultEntryEventArgs<T> : EventArgs
    {
        #region Properties

        /// <summary>
        /// Property contains SearchResultEntry of type T.
        /// </summary>
        public T SearchResultEntry
        {
            get;
            private set;
        }

        #endregion

        #region Fields

        #endregion

        #region Constructors

        /// <summary>
        /// Default constructor for instantiating EwsSearchResultEntryEventArgs
        /// of type T.
        /// </summary>
        /// <param name="searchResultEntry">Type that will hold this class.</param>
        public EwsSearchResultEntryEventArgs(T searchResultEntry)
        {
            this.SearchResultEntry = searchResultEntry;
        }

        #endregion

        #region Public Methods

        #endregion

        #region Private Methods

        #endregion

        #region Override Methods

        #endregion
    }
}
