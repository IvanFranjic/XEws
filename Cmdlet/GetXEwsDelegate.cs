using System;
using System.Collections.Generic;
using System.Management.Automation;
using Microsoft.Exchange.WebServices.Data;
using XEws.CmdletAbstract;

namespace XEws.Cmdlet
{
    [Cmdlet(VerbsCommon.Get, "XEwsDelegate")]
    public class GetXEwsDelegate : XEwsDelegateCmdlet
    {


        protected override void ProcessRecord()
        {
            List<XEwsDelegate> delegates = this.GetDelegate(this.GetSessionVariable());

            WriteObject(delegates, true);
        }

        #region Delegate Override parameters

        [Parameter()]
        new private string DelegateEmailAddress
        {
            get;
            set;
        }

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
