namespace XEws.Cmdlet
{
    using System.Management.Automation;
    using Microsoft.Exchange.WebServices.Data;
    using XEws.CmdletAbstract;

    [Cmdlet(VerbsCommon.Remove, "XEwsDelegate")]
    public sealed class RemoveXEwsDelegate : XEwsDelegateCmdlet
    {
        protected override void ProcessRecord()
        {
            XEwsDelegate xewsDelegate = new XEwsDelegate(this.DelegateEmailAddress);

            this.SetDelegate(xewsDelegate, DelegateAction.Delete);
        }

        #region Delegate Override parameters
        
        [Parameter()]
        new private bool ViewPrivateItems
        {
            get;
            set;
        }


        [Parameter()]
        new private bool ReceiveCopyOfMeetings
        {
            get;
            set;
        }

        [Parameter()]
        new private DelegateFolderPermissionLevel CalendarPermission
        {
            get;
            set;
        }


        [Parameter()]
        new private DelegateFolderPermissionLevel InboxPermission
        {
            get;
            set;
        }

        [Parameter()]
        new private DelegateFolderPermissionLevel TaskPermission
        {
            get;
            set;
        }

        [Parameter()]
        new private DelegateFolderPermissionLevel ContactPermission
        {
            get;
            set;
        }

        #endregion
    }
}
