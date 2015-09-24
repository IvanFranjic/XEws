namespace XEws.CmdletAbstract
{
    using System;
    using System.Collections.Generic;
    using System.Management.Automation;
    using Microsoft.Exchange.WebServices.Data;

    public abstract class XEwsDelegateCmdlet : XEwsCmdlet
    {
        private string delegateEmailAddress = String.Empty;
        [Parameter(Mandatory = true, Position = 0)]
        public string DelegateEmailAddress
        {
            get
            {
                return delegateEmailAddress;
            }
            set
            {
                delegateEmailAddress = value;
            }
        }

        private bool viewPrivateItems = false;
        [Parameter(Mandatory = false, Position = 1)]
        public bool ViewPrivateItems
        {
            get
            {
                return viewPrivateItems;
            }
            set
            {
                viewPrivateItems = value;
            }
        }

        private bool receiveCopyOfMeetings = false;
        [Parameter(Mandatory = false, Position = 2)]
        public bool ReceiveCopyOfMeetings
        {
            get
            {
                return receiveCopyOfMeetings;
            }
            set
            {
                receiveCopyOfMeetings = value;
            }
        }        

        private DelegateFolderPermissionLevel calendarPermission = DelegateFolderPermissionLevel.None;
        [Parameter(Mandatory = false, Position = 3)]
        public DelegateFolderPermissionLevel CalendarPermission
        {
            get
            {
                return calendarPermission;
            }
            set
            {
                calendarPermission = value;
            }
        }

        private DelegateFolderPermissionLevel inboxPermission = DelegateFolderPermissionLevel.None;
        [Parameter(Mandatory = false, Position = 4)]
        public DelegateFolderPermissionLevel InboxPermission
        {
            get
            {
                return inboxPermission;
            }
            set
            {
                inboxPermission = value;
            }
        }

        private DelegateFolderPermissionLevel taskPermission = DelegateFolderPermissionLevel.None;
        [Parameter(Mandatory = false, Position = 5)]
        public DelegateFolderPermissionLevel TaskPermission
        {
            get
            {
                return taskPermission;
            }
            set
            {
                taskPermission = value;
            }
        }

        private DelegateFolderPermissionLevel contactPermission = DelegateFolderPermissionLevel.None;
        [Parameter(Mandatory = false, Position = 6)]
        public DelegateFolderPermissionLevel ContactPermission
        {
            get
            {
                return contactPermission;
            }
            set
            {
                contactPermission = value;
            }
        }

        private MeetingRequestsDeliveryScope meetingRequestDeliveryScope = MeetingRequestsDeliveryScope.DelegatesAndMe;        
        [Parameter(Mandatory = false, Position = 7)]
        public MeetingRequestsDeliveryScope MeetingRequestDeliveryScope
        {
            get
            {
                return this.meetingRequestDeliveryScope;
            }
            set
            {
                this.meetingRequestDeliveryScope = value;
            }
        }

        /// <summary>
        /// Methods is returning all delegate associated with currently binded user.
        /// </summary>
        /// <returns>List of delegates</returns>
        internal List<XEwsDelegate> GetDelegate()
        {
            ExchangeService ewsSession = this.GetSessionVariable();

            List<XEwsDelegate> xewsDelegate = new List<XEwsDelegate>();
            string currentBindedMailbox = this.GetBindedMailbox();
            
            Mailbox currentMailbox = new Mailbox(currentBindedMailbox);

            DelegateInformation xewsDelegateInformation = ewsSession.GetDelegates(currentMailbox, true);

            foreach (DelegateUserResponse delegateResponse in xewsDelegateInformation.DelegateUserResponses)
            {
                if (delegateResponse.Result == ServiceResult.Success)
                {
                    xewsDelegate.Add(new XEwsDelegate(delegateResponse));
                }
            }

            return xewsDelegate;            
        }

        /// <summary>
        /// Method is returning requested delegate.
        /// </summary>
        /// <param name="delegateEmailAddress"></param>
        /// <returns></returns>
        internal XEwsDelegate GetDelegate(string delegateEmailAddress)
        {
            ExchangeService ewsSession = this.GetSessionVariable();

            ValidateEmailAddress(delegateEmailAddress);

            List<XEwsDelegate> ewsDelegates = this.GetDelegate();

            foreach (XEwsDelegate ewsDelegate in ewsDelegates)
            {
                if (ewsDelegate.DelegateUserId.ToLower() == delegateEmailAddress.ToLower())
                {
                    return ewsDelegate;
                }
            }

            throw new InvalidOperationException("No delegate found with email address: " + delegateEmailAddress);
        }

        /// <summary>
        /// Method is used for setting, updating and deleting delegate.
        /// </summary>
        /// <param name="xewsDelegate">Delegate object for manipulation.</param>
        /// <param name="delegateAction">Action on delegate (Add, Update, Delete).</param>
        internal void SetDelegate(XEwsDelegate xewsDelegate, DelegateAction delegateAction)
        {
            ExchangeService ewsSession = this.GetSessionVariable();

            this.ValidateEmailAddress(xewsDelegate.DelegateUserId);
            string currentBindedMailbox = this.GetBindedMailbox();

            DelegateUser delegateUser = new DelegateUser(xewsDelegate.DelegateUserId);
            Mailbox currentMailbox = new Mailbox(currentBindedMailbox);

            delegateUser.ReceiveCopiesOfMeetingMessages = xewsDelegate.ReceivesCopyOfMeeting;
            delegateUser.Permissions.CalendarFolderPermissionLevel = xewsDelegate.CalendarFolderPermission;
            delegateUser.Permissions.InboxFolderPermissionLevel = xewsDelegate.InboxFolderPermission;
            delegateUser.Permissions.ContactsFolderPermissionLevel = xewsDelegate.ContactFolderPermission;
            delegateUser.Permissions.TasksFolderPermissionLevel = xewsDelegate.TaskFolderPermission;
            
            switch (delegateAction)
            {
                case DelegateAction.Update:
                    ewsSession.UpdateDelegates(currentMailbox, MeetingRequestsDeliveryScope.DelegatesAndMe, delegateUser);
                    break;

                case DelegateAction.Add:
                    ewsSession.AddDelegates(currentMailbox, MeetingRequestsDeliveryScope.DelegatesAndMe, delegateUser);
                    break;

                case DelegateAction.Delete:
                    ewsSession.RemoveDelegates(currentMailbox, delegateUser.UserId);
                    break;
            }
        }

        /// <summary>
        /// List of allowed delegate actions.
        /// </summary>
        internal enum DelegateAction
        {
            Update,
            Add,
            Delete
        }

        public sealed class XEwsDelegate
        {
            public XEwsDelegate(DelegateUserResponse delegateResponse)
            {
                this.DelegateUserId = delegateResponse.DelegateUser.UserId.PrimarySmtpAddress;
                this.ReceivesCopyOfMeeting = delegateResponse.DelegateUser.ReceiveCopiesOfMeetingMessages;
                this.ViewPrivateItems = delegateResponse.DelegateUser.ViewPrivateItems;
                this.CalendarFolderPermission = delegateResponse.DelegateUser.Permissions.CalendarFolderPermissionLevel;
                this.InboxFolderPermission = delegateResponse.DelegateUser.Permissions.InboxFolderPermissionLevel;
                this.ContactFolderPermission = delegateResponse.DelegateUser.Permissions.ContactsFolderPermissionLevel;
                this.TaskFolderPermission = delegateResponse.DelegateUser.Permissions.TasksFolderPermissionLevel;
            }

            public XEwsDelegate(string delegateEmailAddress, bool receiveCopiesOfMeeting, bool viewPrivateItems, params DelegateFolderPermissionLevel[] delegatePermissionLevel )
            {
                this.DelegateUserId = delegateEmailAddress;
                this.ReceivesCopyOfMeeting = receiveCopiesOfMeeting;
                this.ViewPrivateItems = viewPrivateItems;
                this.CalendarFolderPermission = delegatePermissionLevel[0];
                this.InboxFolderPermission = delegatePermissionLevel[1];
                this.TaskFolderPermission = delegatePermissionLevel[2];
                this.ContactFolderPermission = delegatePermissionLevel[3];
            }

            public XEwsDelegate(string delegateEmailAddress)
            {
                this.DelegateUserId = delegateEmailAddress;
            }


            public string DelegateUserId
            {
                get;
                private set;
            }

            public bool ReceivesCopyOfMeeting
            {
                get;
                internal set;
            }

            public bool ViewPrivateItems
            {
                get;
                internal set;
            }

            public DelegateFolderPermissionLevel CalendarFolderPermission
            {
                get;
                internal set;
            }

            public DelegateFolderPermissionLevel InboxFolderPermission
            {
                get;
                internal set;
            }

            public DelegateFolderPermissionLevel ContactFolderPermission
            {
                get;
                internal set;
            }

            public DelegateFolderPermissionLevel TaskFolderPermission
            {
                get;
                internal set;
            }

        }
    }
}
