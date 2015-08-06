using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using System.Management.Automation;

namespace XEws.CmdletAbstract
{
    public class XEwsItemMoveCmdlet : XEwsCmdlet
    {
        [Parameter(ValueFromPipeline = true)]
        public Item Item
        {
            get;
            set;
        }

        internal Folder destinationFolder;
        [Parameter()]
        public Folder DestinationFolder
        {
            get
            {
                return this.destinationFolder;
            }
            set
            {
                this.destinationFolder = value;
            }
        }

        #region Methods

        internal void MoveItem(Item item, Folder destinationFolder, MoveOperation moveOperation)
        {
            //Item itemToMove = Item.Bind(ewsSession, item.Id);

            switch (moveOperation)
            {
                case MoveOperation.Copy:
                    item.Copy(destinationFolder.Id);
                    break;

                case MoveOperation.Move:
                    item.Move(destinationFolder.Id);
                    break;
            }
        }

        #endregion

        #region Enums

        public enum MoveOperation
        {
            Copy,
            Move
        }

        #endregion
    }
}
