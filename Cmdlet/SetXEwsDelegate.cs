namespace XEws.Cmdlet
{
    using System.Management.Automation;
    using XEws.CmdletAbstract;
    using System.Collections.Generic;

    [Cmdlet(VerbsCommon.Set, "XEwsDelegate")]
    public sealed class SetXEwsDelegate : XEwsDelegateCmdlet
    {
        protected override void ProcessRecord()
        {
            // Check if requested delegate exist on the specified user account
            XEwsDelegate ewsDelegate = this.GetDelegate(this.DelegateEmailAddress);

            // If delegate exist, get all settings provided by caller and update
            // them accordingly.
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
