using XEws.Helper;
using XEws.Model;

namespace XEws.PowerShell.Cmdlet
{
    using System.Management.Automation;
    using Microsoft.Exchange.WebServices.Data;
    using System.Collections.Generic;
    using System;

    [Cmdlet(VerbsCommon.Get, "EwsMasterCategoryConfiguration")]
    public class GetEwsMasterCategoryConfiguration : EwsCmdlet
    {
        [Parameter(Mandatory = true)]
        public Folder CalendarFolder
        {
            get;
            set;
        }

        [Parameter(Mandatory = false)]
        public SwitchParameter RawData { get; set; }

        protected override void ProcessRecord()
        {

            if (this.CalendarFolder is CalendarFolder)
            {
                try
                {
                    CalendarFolder calendarFolder = (CalendarFolder)this.CalendarFolder;
                    UserConfiguration ewsMasterCategoryConfiguration = UserConfiguration.Bind(this.EwsSession, "CategoryList", calendarFolder.Id, UserConfigurationProperties.All);
                    
                    if (this.RawData.IsPresent)
                        this.WriteObject(ewsMasterCategoryConfiguration);
                    else
                    {
                        string categoryListDefinition = EwsHelper.GetStringFromByte(ewsMasterCategoryConfiguration.XmlData);

                        if (string.IsNullOrEmpty(categoryListDefinition))
                            this.WriteObject(ewsMasterCategoryConfiguration);
                        else
                        {
                            EwsCategoryList ewsCategoryList = new EwsCategoryList(categoryListDefinition);
                            this.WriteObject(ewsCategoryList.ReadCategory(), true);
                        }
                    }
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
