namespace XEws.PowerShell
{
    using Microsoft.Exchange.WebServices.Data;

    public static class EwsWellKnownFolder
    {
        #region Properties

        #endregion

        #region Fields

        #endregion

        #region Constructors

        #endregion

        #region Public Methods

        /// <summary>
        /// Retreives well known folder id (Inbox, SentItems...)
        /// </summary>
        /// <param name="wellKnownFolderName">Wellknownfolder</param>
        /// <returns></returns>
        public static FolderId GetWellKnownFolderId(WellKnownFolderName wellKnownFolderName)
        {
            return new FolderId(wellKnownFolderName);
        }

        #endregion

        #region Private Methods

        #endregion

        #region Override Methods

        #endregion
    }
}
