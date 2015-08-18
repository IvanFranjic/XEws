namespace XEws.CmdletAbstract
{
    using Microsoft.Exchange.WebServices.Data;
    using System.Management.Automation;

    public class XEwsItemMoveCmdlet : XEwsCmdlet
    {
        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
        public Item Item
        {
            get;
            set;
        }

        internal Folder destinationFolder;
        [Parameter(Mandatory = true, Position = 1)]
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

        /// <summary>
        /// Method is moving item from one to another folder.
        /// </summary>
        /// <param name="item">Item to move.</param>
        /// <param name="destinationFolder">Destination folder.</param>
        /// <param name="moveOperation">Move operation (Copy or Move).</param>
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

        /// <summary>
        /// List of allowed move operation.
        /// </summary>
        public enum MoveOperation
        {
            Copy,
            Move
        }

        #endregion
    }
}
