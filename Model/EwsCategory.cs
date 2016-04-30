namespace XEws.Model
{
    using System.Xml;
    using XEws.Helper;

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

        
    }
}
