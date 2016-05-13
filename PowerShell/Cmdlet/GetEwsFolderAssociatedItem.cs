namespace XEws.PowerShell.Cmdlet
{
    using XEws.Model;
    using System.Management.Automation;
    using Microsoft.Exchange.WebServices.Data;

    [Cmdlet(VerbsCommon.Get, "EwsFolderAssociatedItem")]
    public class GetEwsFolderAssociatedItem : EwsCmdlet
    {
        [Parameter(Mandatory = true)]
        public EwsFolderAssociateItemType ItemType
        {
            get;
            set;
        }

        [Parameter(Mandatory = true)]
        public Folder Folder
        {
            get;
            set;
        }

        protected override void ProcessRecord()
        {
            EwsFolderAssociatedItem ewsItem = new EwsFolderAssociatedItem();
            UserConfiguration userConfig = ewsItem.GetFolderAssociatedItem(this.EwsSession, this.Folder, this.ItemType);

            this.WriteObject(userConfig);
        }
    }
}
