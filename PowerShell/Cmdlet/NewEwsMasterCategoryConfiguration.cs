namespace XEws.PowerShell.Cmdlet
{
    using XEws.Model;    
    using System.Management.Automation;
    using Microsoft.Exchange.WebServices.Data;
    using CalendarFolder = Microsoft.Exchange.WebServices.Data.CalendarFolder;

    [Cmdlet(VerbsCommon.New, "EwsMasterCategoryConfiguration")]
    public class NewEwsMasterCategoryConfiguration : EwsCmdlet
    {
        #region Properties

        [Parameter(Mandatory = true)]
        public Folder CalendarFolder
        {
            get;
            set;
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
            if (!(this.CalendarFolder is CalendarFolder))
                this.WriteWarning("Provided folder is not Calendar folder. Category list can be only stored in Calendar folder.");

            else
            {
                try
                {
                    CalendarFolder calendarFolder = (CalendarFolder)this.CalendarFolder;
                    UserConfiguration ewsMasterCategoryConfiguration = UserConfiguration.Bind(this.EwsSession, "CategoryList", calendarFolder.Id, UserConfigurationProperties.All);
                    this.WriteWarning("CategoryList already exist.");
                    return;
                }
                catch
                {
                    EwsCategory.RebuildCategoryList(this.EwsSession, CalendarFolder.Id);
                }
            }
            
        }

        #endregion
    }
}
