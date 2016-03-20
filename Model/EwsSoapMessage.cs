namespace XEws.Model
{
    using System.IO;
    using System.Xml;
    using Microsoft.Exchange.WebServices.Data;

    public abstract class EwsSoapMessage
    {
        // TODO: create method that will read mail tip from soap message
        // xml reader...
        internal void ReadMailTipResponse()
        {

        }
        
        // TODO: create method that will create mail tip request
        internal void CreateMailTipRequest(string sender, string recipient, EwsMailTipType mailTipType)
        {


        }
    }
}
