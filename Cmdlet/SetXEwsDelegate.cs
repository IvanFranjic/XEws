namespace XEws.Cmdlet
{
    using System;
    using System.Management.Automation;
    using Microsoft.Exchange.WebServices.Data;
    using XEws.CmdletAbstract;
    using System.Collections.Generic;

    [Cmdlet(VerbsCommon.Set, "XEwsDelegate")]
    public class SetXEwsDelegate : XEwsDelegateCmdlet
    {
        protected override void ProcessRecord()
        {
            XEwsDelegate ewsDelegate = this.GetDelegate(this.DelegateEmailAddress);

            foreach (KeyValuePair<string, object> passedParameter in this.MyInvocation.BoundParameters)
            {
                switch (passedParameter.Key.ToString())
                {
                    case "CalendarPermission":
                        ewsDelegate.CalendarFolderPermission = this.CalendarPermission;
                        break;

                    case "InboxPermission":
                        ewsDelegate.InboxFolderPermission = this.InboxPermission;
                        break;

                    case "TaskPermission":
                        ewsDelegate.TaskFolderPermission = this.TaskPermission;
                        break;

                    case "ContactPermission":
                        ewsDelegate.ContactFolderPermission = this.ContactPermission;
                        break;

                    case "ViewPrivateItems":
                        ewsDelegate.ViewPrivateItems = this.ViewPrivateItems;
                        break;

                    case "ReceiveCopyOfMeetings":
                        ewsDelegate.ReceivesCopyOfMeeting = this.ReceiveCopyOfMeetings;
                        break;
                }
            }

            this.SetDelegate(ewsDelegate, DelegateAction.Update);
        }
    }
}
