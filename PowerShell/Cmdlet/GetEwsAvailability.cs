namespace XEws.PowerShell.Cmdlet
{
    using System;
    using System.Collections.Generic;
    using System.Management.Automation;
    using Microsoft.Exchange.WebServices.Data;

    [Cmdlet(VerbsCommon.Get, "EwsAvailability")]
    public class GetEwsAvailability : EwsCmdlet
    {
        #region Properties

        private string emailAddress = null;
        [Parameter(Mandatory = true, Position = 0)]
        public string EmailAddress
        {
            get
            {
                return emailAddress;
            }
            set
            {
                this.ValidateEmailAddress(value.ToString());
                this.emailAddress = value;
            }
        }

        private DateTime startTime = DateTime.Now;
        [Parameter(Mandatory = false, Position = 1)]
        public DateTime StartTime
        {
            get
            {
                return startTime;
            }
            set
            {
                this.startTime = value;
            }
        }

        private DateTime endTime = DateTime.Now.AddDays(1);
        [Parameter(Mandatory = false, Position = 1)]
        public DateTime EndTime
        {
            get
            {
                return endTime;
            }
            set
            {
                this.endTime = value;
            }
        }        

        #endregion

        #region Fields

        #endregion

        #region Constructors

        #endregion

        #region Public Methods

        #endregion

        #region Private Methods

        #endregion

        #region Override Methods

        protected override void ProcessRecord()
        {
            AttendeeInfo attendee = new AttendeeInfo(this.EmailAddress, MeetingAttendeeType.Required, false);
            List<AttendeeInfo> attendeeList = new List<AttendeeInfo>();
            attendeeList.Add(attendee);

            AvailabilityOptions options = new AvailabilityOptions();
            options.MeetingDuration = 30;
            options.RequestedFreeBusyView = FreeBusyViewType.FreeBusy;

            GetUserAvailabilityResults freeBusyAvailability = this.EwsSession.GetUserAvailability(
                attendeeList,
                new TimeWindow(this.StartTime, this.EndTime),
                AvailabilityData.FreeBusy,
                options
            );

            foreach ( AttendeeAvailability availability in freeBusyAvailability.AttendeesAvailability )
            {
                if (availability.Result == ServiceResult.Success)
                    this.WriteObject(availability.CalendarEvents, true);
                else
                    this.WriteObject(availability);
            }
        }

        #endregion

    }
}
