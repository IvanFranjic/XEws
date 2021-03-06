﻿namespace XEws.Tracer
{
    using System;
    using Microsoft.Exchange.WebServices.Data;
    using System.Xml;
    using System.IO;

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
