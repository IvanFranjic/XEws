using System;
using System.Collections.Generic;
using System.Text;
using XEws.Helper;

namespace XEws.PowerShell.Cmdlet
{
    using XEws.Model;
    using System.Management.Automation;
    using Microsoft.Exchange.WebServices.Data;
    using CalendarFolder = Microsoft.Exchange.WebServices.Data.CalendarFolder; 

    [Cmdlet(VerbsData.Update, "EwsMasterCategoryConfiguration")]
    public class UpdateEwsMasterCategoryConfiguration : EwsCmdlet
    {
        #region Properties

        [Parameter(Mandatory = true)]
        public Folder CalendarFolder
        {
            get;
            set;
        }

        [Parameter(Mandatory = true)]
        public List<EwsCategory> Categories { get; set; }

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

                    string categoryListDefinition = EwsHelper.GetStringFromByte(ewsMasterCategoryConfiguration.XmlData);

                    if (string.IsNullOrEmpty(categoryListDefinition))
                    {
                        StringBuilder xmlBuilder = new StringBuilder();
                        xmlBuilder.AppendLine("<?xml version=\"1.0\"?>");
                        xmlBuilder.AppendLine("<categories default=\"\" lastSavedSession=\"\" lastSavedTime=\"\" xmlns=\"CategoryList.xsd\">");
                        xmlBuilder.AppendLine("</categories>");

                        categoryListDefinition = xmlBuilder.ToString();
                    }
                    
                    EwsCategoryList ewsCategoryList = new EwsCategoryList(categoryListDefinition);
                    List<EwsCategory> currentCategories = ewsCategoryList.ReadCategory();

                    foreach (EwsCategory category in currentCategories)
                    {
                        if (this.Categories.Contains(category))
                            this.Categories.Remove(category);
                    }

                    string newCategories = ewsCategoryList.InsertElement(this.Categories);
                    
                    byte[] xmlData = Encoding.UTF8.GetBytes(newCategories);

                    EwsCategory.UpdateCategoryList(this.EwsSession, this.CalendarFolder.Id, xmlData);

                }
                catch (Exception e)
                {
                    throw new InvalidOperationException(e.Message);
                }
            }
        }

        #endregion
    }
}
