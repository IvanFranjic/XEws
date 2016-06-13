namespace XEws.Model
{
    using System;
    using System.Xml;
    using XEws.Helper;
    using System.Reflection;
    using Microsoft.Exchange.WebServices.Data;

    //public class EwsCategory
    //{
    //    public string CategoryName { get; private set; }
    //    public int CategoryColor { get; private set; }
    //    public string CategoryGuid { get; private set; }


    //    public EwsCategory(string categoryName, string categoryGuid, int categoryColor)
    //    {
    //        this.CategoryName = categoryName;
    //        this.CategoryGuid = categoryGuid;
    //        this.CategoryColor = categoryColor;
    //    }

    //    public EwsCategory(byte[] categoryByteData)
    //    {
    //        string categoryData = EwsHelper.GetStringFromByte(categoryByteData);

    //    }

    //    public static void RebuildCategoryList(ExchangeService ewsSession, FolderId calendarFolderId)
    //    {
    //        UserConfiguration userConfig = new UserConfiguration(ewsSession);
    //        userConfig.Save("CategoryList", calendarFolderId);
    //    }
        
    //}

    public struct EwsCategory
    {
        public int UsageCount { get; private set; }

        public int LastSessionUsed { get; private set; }

        public int RenameOnFirstUse { get; private set; }

        public string Name { get; private set; }

        public int Color { get; private set; }

        public string KeyboardShortcut { get; private set; }

        public DateTime LastTimeUsedNotes { get; private set; }

        public DateTime LastTimeUsedJournal { get; private set; }

        public DateTime LastTimeUsedContacts { get; private set; }

        public DateTime LastTimeUsedTasks { get; private set; }

        public DateTime LastTimeUsedCalendar { get; private set; }
        public DateTime LastTimeUsedMail { get; private set; }

        public DateTime LastTimeUsed { get; private set; }

        public string Guid { get; private set; }

        public EwsCategory(int usageCount, int lastSessionUsed, int renameOnFirstUse, string name, int color, string keyBoardShortCut, DateTime lastTimeUsedNotes, DateTime lastTimeUsedJournal, DateTime lastTimeUsedContacts, DateTime lastTimeUsedTasks,
            DateTime lastTimeUsedCalendar, DateTime lastTimeUsedMail, DateTime lastTimeUsed, string guid)
        {
            UsageCount = usageCount;
            LastSessionUsed = lastSessionUsed;
            RenameOnFirstUse = renameOnFirstUse;
            Name = name;
            Color = color;
            KeyboardShortcut = keyBoardShortCut;
            LastTimeUsedNotes = lastTimeUsedNotes;
            LastTimeUsedJournal = lastTimeUsedJournal;
            LastTimeUsedContacts = lastTimeUsedContacts;
            LastTimeUsedTasks = lastTimeUsedTasks;
            LastTimeUsedCalendar = lastTimeUsedCalendar;
            LastTimeUsedMail = lastTimeUsedMail;
            LastTimeUsed = lastTimeUsed;

            if (!guid.StartsWith("{") && !guid.EndsWith("}"))
                Guid = string.Format("{{{0}}}", guid.ToUpper());
            else
                Guid = guid.ToUpper();
        }

        public EwsCategory(XmlNode categoryNode) :
            this(
                    Convert.ToInt32(categoryNode.Attributes["usageCount"]?.InnerText),
                    Convert.ToInt32(categoryNode.Attributes["lastSessionUsed"]?.InnerText),
                    Convert.ToInt32(categoryNode.Attributes["renameOnFirstUse"]?.InnerText),
                    categoryNode.Attributes["name"]?.InnerText,
                    Convert.ToInt32(categoryNode.Attributes["color"]?.InnerText),
                    categoryNode.Attributes["keyboardShortcut"]?.InnerText,
                    Convert.ToDateTime(categoryNode.Attributes["lastTimeUsedNotes"]?.InnerText),
                    Convert.ToDateTime(categoryNode.Attributes["lastTimeUsedJournal"]?.InnerText),
                    Convert.ToDateTime(categoryNode.Attributes["lastTimeUsedContacts"]?.InnerText),
                    Convert.ToDateTime(categoryNode.Attributes["lastTimeUsedTasks"]?.InnerText),
                    Convert.ToDateTime(categoryNode.Attributes["lastTimeUsedCalendar"]?.InnerText),
                    Convert.ToDateTime(categoryNode.Attributes["lastTimeUsedMail"]?.InnerText),
                    Convert.ToDateTime(categoryNode.Attributes["lastTimeUsed"]?.InnerText),
                    categoryNode.Attributes["guid"]?.InnerText
                )
        {
        }

        public EwsCategory(int color, string categoryName)
            : this(0, 0, 0, categoryName, color, "0", DateTime.Now,
                  DateTime.Now, DateTime.Now, DateTime.Now, DateTime.Now,
                  DateTime.Now, DateTime.Now, System.Guid.NewGuid().ToString())
        {
            if (color > 10 || color < 0)
                color = 10;

            Color = color;
        }

        public override string ToString()
        {
            return this.Name;
        }

        public override bool Equals(object ewsCategory)
        {
            if (!(ewsCategory is EwsCategory))
                return false;

            EwsCategory cat = (EwsCategory)ewsCategory;

            return ((this.Name == cat.Name) &&
                    (this.Color == cat.Color));
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        public string GetPropertyValue(string propertyName)
        {
            string propName = string.Format("{0}{1}",
                propertyName.Substring(0, 1).ToUpper(),
                propertyName.Substring(1)
            );

            PropertyInfo propertyInfo =
                this.GetType().GetProperty(propName);

            return propertyInfo?.GetValue(this).ToString();
        }

        public static void RebuildCategoryList(ExchangeService ewsSession, FolderId calendarFolderId, byte[] xmlData)
        {
            UserConfiguration userConfig = new UserConfiguration(ewsSession);
            userConfig.XmlData = xmlData;
            userConfig.Save("CategoryList", calendarFolderId);
        }

        public static void UpdateCategoryList(ExchangeService ewsSession, FolderId calendarFolderId, byte[] xmlData)
        {
            UserConfiguration userConfig = UserConfiguration.Bind(ewsSession, "CategoryList", calendarFolderId, UserConfigurationProperties.All);
            userConfig.XmlData = xmlData;

            if (userConfig.IsDirty)
                userConfig.Update();
        }
    }
}
