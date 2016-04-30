namespace XEws.Model
{
    using System.Xml;
    using XEws.Helper;
    using Microsoft.Exchange.WebServices.Data;

    public class EwsCategory
    {
        public string CategoryName { get; private set; }
        public int CategoryColor { get; private set; }
        public string CategoryGuid { get; private set; }


        public EwsCategory(string categoryName, string categoryGuid, int categoryColor)
        {
            this.CategoryName = categoryName;
            this.CategoryGuid = categoryGuid;
            this.CategoryColor = categoryColor;
        }

        public EwsCategory(byte[] categoryByteData)
        {
            string categoryData = EwsHelper.GetStringFromByte(categoryByteData);

        }

        public static void RebuildCategoryList(ExchangeService ewsSession, FolderId calendarFolderId)
        {
            UserConfiguration userConfig = new UserConfiguration(ewsSession);
            userConfig.Save("CategoryList", calendarFolderId);
        }
        
    }
}
