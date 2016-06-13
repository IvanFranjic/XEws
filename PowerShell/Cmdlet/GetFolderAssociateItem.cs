using Microsoft.Exchange.WebServices.Data;
using XEws.Model;

namespace XEws.PowerShell.Cmdlet
{
    using System.Management.Automation;

    [Cmdlet(VerbsCommon.Get, "FolderAssociateItem")]
    public class GetFolderAssociateItem : EwsCmdlet
    {
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        public EwsFolderAssociateItemType FolderAssociateItemType { get; set; }

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        public WellKnownFolderName FolderName { get; set; }

        protected override void ProcessRecord()
        {
            EwsFolderAssociatedItem ewsAssociatedItem = new EwsFolderAssociatedItem();
            UserConfiguration userConfig = ewsAssociatedItem.GetFolderAssociatedItem(this.EwsSession, FolderName, FolderAssociateItemType);
            WriteObject(userConfig);
        }
    }
}
