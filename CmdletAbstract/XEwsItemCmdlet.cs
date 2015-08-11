using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using Microsoft.Exchange.WebServices.Data;
using System.Dynamic;

namespace XEws.CmdletAbstract
{
    public abstract class XEwsItemCmdlet : XEwsCmdlet
    {
        [Parameter()]
        public Folder Folder
        {
            get;
            set;
        }

        #region Methods

        internal List<Item> GetItem(Folder folder, ItemView iView)
        {
            FindItemsResults<Item> findResult = folder.FindItems(iView);
            return this.GetListItem(findResult);
        }

        internal List<Item> GetItem(Folder folder, ItemView iView, DateTime startDate, DateTime endDate)
        {
            //SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(
            //    LogicalOperator.And,
            //    new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived, startDate),
            //    new SearchFilter.IsLessThanOrEqualTo(ItemSchema.DateTimeReceived, endDate)
            //    );

            //FindItemsResults<Item> findResult = folder.FindItems(searchFilter, iView);

            //return this.GetListItem(findResult);

            return this.GetItem(folder, iView, startDate, endDate, String.Empty);
        }

        internal List<Item> GetItem(Folder folder, ItemView iView, DateTime startDate, DateTime endDate, string subjectToSearch)
        {
            SearchFilter.SearchFilterCollection searchFilterCollection = new SearchFilter.SearchFilterCollection(LogicalOperator.And);
            searchFilterCollection.Add(new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived, startDate));
            searchFilterCollection.Add(new SearchFilter.IsLessThanOrEqualTo(ItemSchema.DateTimeReceived, endDate));

            if (!String.IsNullOrEmpty(subjectToSearch))
            {
                searchFilterCollection.Add(new SearchFilter.IsEqualTo(ItemSchema.Subject, subjectToSearch));
            }

            SearchFilter searchFilter = searchFilterCollection;
            FindItemsResults<Item> findResult = folder.FindItems(searchFilter, iView);

            return this.GetListItem(findResult);
        }

        private List<Item> GetListItem(FindItemsResults<Item> findResult)
        {
            List<Item> items = new List<Item>();
            foreach (Item item in findResult)
                items.Add(item);

            return items;
        }

        #endregion

        

    }

}
