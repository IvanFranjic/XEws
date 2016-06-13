namespace XEws.Model
{
    using System;
    using System.Xml;
    using System.Collections.Generic;

    public class EwsCategoryList : EwsFolderAssociatedItem
    {
        private string xmlStringData;

        public EwsCategoryList(string loadFilePath)
        {
            xmlStringData = loadFilePath;
        }


        public List<EwsCategory> ReadCategory()
        {
            List<EwsCategory> oneCategories = new List<EwsCategory>();
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.LoadXml(xmlStringData);
            try
            {
                foreach (XmlNode childs in xmlDocument.ChildNodes)
                {
                    if (!xmlDocument.HasChildNodes) continue;

                    foreach (XmlNode childsFromChild in childs)
                        if (childsFromChild.Attributes != null)
                            oneCategories.Add(new EwsCategory(childsFromChild));
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            return oneCategories;
        }

        public string InsertElement(EwsCategory category)
        {
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.LoadXml(xmlStringData);
            XmlElement categoryElement = xmlDocument.CreateElement("category", "CategoryList.xsd"); //<category/>

            string[] xmlAttributes = new[]
            {
                "usageCount", "lastSessionUsed", "renameOnFirstUse",
                "name", "color", "keyboardShortcut", "lastTimeUsedNotes",
                "lastTimeUsedJournal", "lastTimeUsedContacts", "lastTimeUsedTasks",
                "lastTimeUsedCalendar", "lastTimeUsedMail", "lastTimeUsed", "guid"
            };

            foreach (string xmlAttribute in xmlAttributes)
            {
                XmlAttribute attribute = xmlDocument.CreateAttribute(xmlAttribute);
                attribute.Value = category.GetPropertyValue(xmlAttribute);
                categoryElement.Attributes.Append(attribute);
            }

            xmlDocument.DocumentElement?.AppendChild(categoryElement);

            return xmlDocument.OuterXml;
        }

        public string InsertElement(List<EwsCategory> categories)
        {
            foreach (EwsCategory category in categories)
            {
                xmlStringData = this.InsertElement(category);
            }

            return xmlStringData;
        }
    }
}
