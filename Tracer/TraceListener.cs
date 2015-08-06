using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using System.Xml;
using System.IO;

namespace XEws.Tracer
{
    public class TraceListener : ITraceListener
    {
        public TraceListener(string tracePath)
        {
            this.TracePath = tracePath;
        }

        internal string TracePath
        {
            get;
            private set;
        }

        #region ITraceListener Members

        public void Trace(string traceType, string traceMessage)
        {
            CreateXMLTextFile(traceType, traceMessage);
        }

        #endregion

        private void CreateXMLTextFile(string fileName, string traceContent)
        {
            string tracePath = String.Format("{0}\\{1}", this.TracePath, fileName);

            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(traceContent);
                xmlDoc.Save(tracePath + ".xml");

                //this.ResponseVariable = traceContent;
            }
            catch (Exception)
            {
                File.WriteAllText(tracePath + ".txt", traceContent);
            }
        }
    }
}
