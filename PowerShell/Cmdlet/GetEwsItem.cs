namespace XEws.PowerShell.Cmdlet
{
    using System;
    using XEws.Model;
    using System.Management.Automation;
    using Microsoft.Exchange.WebServices.Data;

    [Cmdlet(VerbsCommon.Get, "EwsItem")]
    public class GetEwsItem : EwsCmdletAbstract
    {
        #region Properties

        private FolderId folderRoot = EwsWellKnownFolder.GetWellKnownFolderId(WellKnownFolderName.Inbox);
        [Parameter(Position = 0)]
        public FolderId FolderRoot
        {
            get
            {
                return this.folderRoot;
            }
            set
            {
                this.folderRoot = value;
            }
        }

        private DateTime startDate = DateTime.Now.AddDays(-15);
        [Parameter(Position = 1)]
        public DateTime StartDate
        {
            get
            {
                return this.startDate;
            }
            set
            {
                this.startDate = value;
            }
        }

        private DateTime endDate = DateTime.Now;
        [Parameter(Position = 2)]
        public DateTime EndDate
        {
            get
            {
                return this.endDate;
            }
            set
            {
                this.endDate = value;
            }
        }

        private SearchFilter searchFilter;
        [Parameter(Position = 3)]
        public SearchFilter SearchFilter
        {
            // if filter is not provided, instantiate default one
            get
            {
                if (this.searchFilter != null)
                    return this.searchFilter;

                this.searchFilter = this.CreateDefaultSearchFilter();

                return this.searchFilter;
            }
            set
            {
                this.searchFilter = new SearchFilter.SearchFilterCollection(
                    LogicalOperator.And,
                    this.CreateDefaultSearchFilter(),
                    value
                );
            }
        }

        private PropertyDefinitionBase[] customProperties = null;
        [Parameter()]
        public PropertyDefinitionBase[] CustomProperty
        {
            get
            {
                return this.customProperties;
            }
            set
            {
                this.customProperties = value;
            }
        }

        #endregion

        #region Fields

        #endregion

        #region Constructors

        #endregion

        #region Public Methods

        #endregion

        #region Private Methods

        /// <summary>
        /// Instantiate default search filter with start and end date
        /// for item.
        /// </summary>
        /// <returns>SearchFilter</returns>
        private SearchFilter CreateDefaultSearchFilter()
        {
            SearchFilter defaultSearchFilter = new SearchFilter.SearchFilterCollection(
                    LogicalOperator.And,
                    new SearchFilter.IsGreaterThan(ItemSchema.DateTimeCreated, this.startDate),
                    new SearchFilter.IsLessThan(ItemSchema.DateTimeCreated, this.endDate)
                );

            return defaultSearchFilter;
        }

        #endregion

        #region Override Methods

        protected override void ProcessRecord()
        {
            ItemView itemView = new ItemView(30);

            if (this.CustomProperty != null)
            {
                PropertySet propertySet = new PropertySet(BasePropertySet.FirstClassProperties);

                foreach (PropertyDefinitionBase customProp in this.CustomProperty)
                    propertySet.Add(customProp);

                itemView.PropertySet = propertySet;
            }            
            
            EwsItem ewsItems = new EwsItem(this.EwsSession, this.FolderRoot, itemView, this.SearchFilter);
            ewsItems.SearchResultFound += OnSearchResultFound;
            ewsItems.FindItem(this.FolderRoot);
        }

        /// <summary>
        /// This will be subscribed to SearchResultFound event.
        /// </summary>
        /// <param name="sender">Sender of the event.</param>
        /// <param name="e">EwsSearchResultEntryEventArgs<Item></Item></param>
        private void OnSearchResultFound(object sender, EwsSearchResultEntryEventArgs<Item> e)
        {
            // If we are subscribed to event, output object to the PS console.
            this.WriteObject(e.SearchResultEntry, true);
        }

        #endregion
    }
}
