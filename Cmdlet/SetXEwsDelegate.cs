using System;
using System.Management.Automation;
using Microsoft.Exchange.WebServices.Data;
using XEws.CmdletAbstract;

namespace XEws.Cmdlet
{
    [Cmdlet(VerbsCommon.Set, "XEwsDelegate")]
    public class SetXEwsDelegate : XEwsDelegateCmdlet
    {
        protected override void ProcessRecord()
        {
            XEwsDelegate ewsDelegate = this.GetDelegate(this.DelegateEmailAddress);

            ewsDelegate.CalendarFolderPermission = this.CalendarPermission;
            ewsDelegate.InboxFolderPermission = this.InboxPermission;
            ewsDelegate.TaskFolderPermission = this.TaskPermission;
            ewsDelegate.ContactFolderPermission = this.ContactPermission;
            ewsDelegate.ViewPrivateItems = this.ViewPrivateItems;
            ewsDelegate.ReceivesCopyOfMeeting = this.ReceiveCopyOfMeetings;

            this.SetDelegate(ewsDelegate, DelegateAction.Update);

            //DelegateFolderPermissionLevel[] folderPermission = { this.CalendarPermission, this.InboxPermission, this.TaskPermission, this.ContactPermission };
            //XEwsDelegate delegateUser = new XEwsDelegate(this.DelegateEmailAddress, this.ReceiveCopyOfMeetings, this.ViewPrivateItems, folderPermission);

            //this.AddDelegate(delegateUser, this.GetSessionVariable());
        }
    }
}
