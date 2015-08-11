using System;
using System.Management.Automation;
using Microsoft.Exchange.WebServices.Data;
using XEws.CmdletAbstract;

namespace XEws.Cmdlet
{
    [Cmdlet(VerbsCommon.Add, "XEwsDelegate")]
    public class AddXEwsDelegate : XEwsDelegateCmdlet
    {
        protected override void ProcessRecord()
        {
            DelegateFolderPermissionLevel[] folderPermission = { this.CalendarPermission, this.InboxPermission, this.TaskPermission, this.ContactPermission };
            XEwsDelegate delegateUser = new XEwsDelegate(this.DelegateEmailAddress, this.ReceiveCopyOfMeetings, this.ViewPrivateItems, folderPermission);
            
            this.SetDelegate(delegateUser, DelegateAction.Add,this.GetSessionVariable());
        }
    }
}
