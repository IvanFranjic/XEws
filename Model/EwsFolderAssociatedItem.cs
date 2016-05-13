namespace XEws.Model
{
    using System;
    using Microsoft.Exchange.WebServices.Data;

    public class EwsFolderAssociatedItem
    {

        /// <summary>
        /// Method will verify if FAI is part of that folder. 
        /// Implementation should be simple. Something like: 
        /// if ItemType should reside in Folder return true, otherwise false.
        /// For example CategoryList is only stored in Calendar root folder
        /// either in Archive or Primary mailbox.
        /// </summary>
        /// <param name="folderAssociatedItem">Type of folder associated item</param>
        /// <param name="storeFolder"></param>
        /// <returns></returns>
        internal virtual bool VerifyItemFolderAssociation(EwsFolderAssociateItemType folderAssociatedItem, Folder storeFolder)
        {
            switch (folderAssociatedItem)
            {
                // Category list is stored in calendar
                case EwsFolderAssociateItemType.CategoryList:
                    // TODO: check if it's root calendar.
                    if (storeFolder is CalendarFolder)
                        return true;
                    break;

                case EwsFolderAssociateItemType.AvailabilityOptions:
                    // TODO: check if it's root calendar.
                    if (storeFolder is CalendarFolder)
                        return true;
                    break;
            }

            // If we get to this point just return false since we
            // are not aware of that item.
            return false;
        }

        public virtual UserConfiguration GetFolderAssociatedItem(ExchangeService ewsSession, Folder faiFolder, EwsFolderAssociateItemType itemType)
        {
            if (!this.VerifyItemFolderAssociation(itemType, faiFolder))
            {
                throw new InvalidOperationException(
                    string.Format("Item '{0}' is not expected to be located in folder '{1}'", itemType.ToString(), faiFolder.DisplayName)    
                );
            }

            UserConfiguration userConfig = UserConfiguration.Bind(ewsSession, itemType.ToString(), faiFolder.Id, UserConfigurationProperties.All);

            return userConfig;
        }
    }
}
