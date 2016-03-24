namespace XEws.Model
{
    using System.IO;
    using System.Xml;
    using System.Net;
    using Microsoft.Exchange.WebServices.Data;

    public abstract class EwsSoapMessage
    {
        internal XmlDocument ReadSoapDocument(string soapData)
        {
            XmlDocument soapDocument = new XmlDocument();
            try
            {
                soapDocument.Load(soapData);
            }
            catch
            {
                // TODO: for now just supress errors.
            }

            return soapDocument;
        }

        internal WebRequest CreateSoapWebRequest(string soapData, string uri, ICredentials credentials)
        {
            WebRequest soapWebRequest = WebRequest.Create(uri);
            soapWebRequest.Headers.Set(HttpRequestHeader.Pragma, "no-cache");
            soapWebRequest.Headers.Set(HttpRequestHeader.Translate, "f");
            soapWebRequest.ContentType = "text/xml";
            soapWebRequest.ContentLength = soapData.Length;
            soapWebRequest.Timeout = 60000;
            soapWebRequest.Method = "POST";
            soapWebRequest.Credentials = credentials;

            return soapWebRequest;

        }

        internal string GetSoapWebResponse(WebRequest webRequest)
        {
            WebResponse webResponse = webRequest.GetResponse();
            StreamReader streamReader = new StreamReader(
                webResponse.GetResponseStream()
            );

            return streamReader.ReadToEnd();
        }
    }
}
