namespace XEws.CmdletAbstract
{
    using Microsoft.Exchange.WebServices.Data;
    using System.Collections.Generic;

    public class XEwsContactCmdlet : XEwsCmdlet
    {
        internal List<Contact> GetContact(string something)
        {
            ExchangeService ewsSession = this.GetSessionVariable();
            List<Contact> contacts = new List<Contact>();
            Folder contactFolder = XEwsFolderCmdlet.GetWellKnownFolder(WellKnownFolderName.Contacts, ewsSession);

            int numberOfContacts = 50;
            ItemView iView = new ItemView(numberOfContacts);
            iView.PropertySet = new PropertySet(BasePropertySet.IdOnly, ContactSchema.DisplayName);

            FindItemsResults<Item> findContacts;

            do
            {
                findContacts = ewsSession.FindItems(contactFolder.Id, iView);

                foreach (Item item in findContacts)
                {
                    if (item is Contact)
                    {
                        Contact contact = item as Contact;
                        contacts.Add(contact);
                    }
                }

                iView.Offset += 50;

            } while (findContacts.MoreAvailable);

            return contacts;
        }

        internal void GetContact()
        {
            ExchangeService ewsSession = this.GetSessionVariable();
            NameResolutionCollection addressCollection = ewsSession.ResolveName("SMTP:", ResolveNameSearchLocation.DirectoryOnly, true);

            foreach (NameResolution address in addressCollection)
            {
                WriteObject(address);
            }
        }
    }

    public sealed class XEwsContact
    {
        public string DisplayName
        {
            get;
            set;
        }

        public string EmailAddress
        {
            get;
            set;
        }
        
        public XEwsContact(Contact contact)
        {
            this.DisplayName = contact.DisplayName;
            // this.EmailAddress = contact.EmailAddresses[EmailAddressKey.EmailAddress1].ToString();
        }
    }
}
