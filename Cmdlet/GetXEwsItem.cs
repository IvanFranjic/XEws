using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using Microsoft.Exchange.WebServices.Data;
using System.Collections;
using XEws.CmdletAbstract;

namespace XEws.Cmdlet
{
    [Cmdlet(VerbsCommon.Get, "XEwsItem")]
    public class GetXEwsItem : XEwsItemCmdlet
    {
        private int itemsToReturn = 100;
        [Parameter()]
        public int ItemsToReturn
        {
            get
            {
                return itemsToReturn;
            }
            set
            {
                itemsToReturn = value;
            }
        }

        private DateTime startDate = DateTime.Now.AddYears(-5);
        [Parameter()]
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
        [Parameter()]
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

        private string messageSubject = String.Empty;
        [Parameter()]
        public string MessageSubject
        {
            get
            {
                return this.messageSubject;
            }
            set
            {
                this.messageSubject = value;
            }
        }

        protected override void ProcessRecord()
        {
            if (this.Folder == null)
                this.Folder = XEwsFolderCmdlet.GetWellKnownFolder(WellKnownFolderName.Inbox, this.GetSessionVariable());

            int pageSize = 50;
            int remainingMessages = this.itemsToReturn;

            ItemView iView = new ItemView(pageSize);

            for (int i = 0; i < this.itemsToReturn; i += pageSize)
            {                
                iView.Offset = i;                

                if (remainingMessages < pageSize)
                {
                    pageSize = remainingMessages;
                    iView.PageSize = pageSize;
                }

                remainingMessages -= pageSize;

                List<Item> items = this.GetItem(this.Folder, iView, startDate, endDate, messageSubject);
                WriteObject(items, true);
            }
        }        
    }
}
