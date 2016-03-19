namespace XEws.Cmdlet
{
    using System.Management.Automation;
    using Microsoft.Exchange.WebServices.Data;
    using XEws.CmdletAbstract;

    // Removing command from the module, new set of Ews command will be
    // introduced.
    //[Cmdlet(VerbsCommon.Add, "XEwsDelegate")]
    public sealed class AddXEwsDelegate : XEwsDelegateCmdlet
    {
        protected override void ProcessRecord()
        {
            DelegateFolderPermissionLevel[] folderPermission = { this.CalendarPermission, this.InboxPermission, this.TaskPermission, this.ContactPermission };
            XEwsDelegate delegateUser = new XEwsDelegate(this.DelegateEmailAddress, this.ReceiveCopyOfMeetings, this.ViewPrivateItems, folderPermission);
            
            this.SetDelegate(delegateUser, DelegateAction.Add);
        }
    }
}
