namespace XEws.PowerShell.Cmdlet
{
    using XEws.Model;    
    using System.Management.Automation;
    using Microsoft.Exchange.WebServices.Data;

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
            EwsCategory.RebuildCategoryList(this.EwsSession, CalendarFolder.Id);
        }

        #endregion
    }
}
