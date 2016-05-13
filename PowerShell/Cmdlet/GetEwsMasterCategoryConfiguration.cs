namespace XEws.PowerShell.Cmdlet
{
    using System.Management.Automation;
    using Microsoft.Exchange.WebServices.Data;
    using System.Collections.Generic;
    using System;

    [Cmdlet(VerbsCommon.Get, "EwsMasterCategoryConfiguration")]
    public class GetEwsMasterCategoryConfiguration : EwsCmdlet
    {
        [Parameter(Mandatory = false)]
        public Folder CalendarFolder
        {
            get;
            set;
        }

        protected override void ProcessRecord()
        {

            if (this.CalendarFolder is CalendarFolder)
            {
                try
                {
                    CalendarFolder calendarFolder = (CalendarFolder)this.CalendarFolder;
                    UserConfiguration ewsMasterCategoryConfiguration = UserConfiguration.Bind(this.EwsSession, "CategoryList", calendarFolder.Id, UserConfigurationProperties.All);

                    this.WriteObject(ewsMasterCategoryConfiguration);
                }
                catch (ServiceResponseException responseException)
                {
                    if (responseException.ErrorCode == ServiceError.ErrorItemNotFound)
                        this.WriteWarning("CategoryList item not found in specified calendar folder.");
                    else
                        this.WriteWarning(responseException.Message);

                }
                catch (Exception exception)
                {
                    throw new InvalidOperationException(exception.Message);
                }
            }
            else
                this.WriteWarning("Provided folder is not calendar folder. Please specify calendar folder since Category list is stored in Calendar.");            
        }
    }
}
