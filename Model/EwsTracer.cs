namespace XEws.Model
{
    using System;
    using System.IO;
    using Microsoft.Exchange.WebServices.Data;
    
    public class EwsTracer : ITraceListener
    {
        #region Properties

        #endregion

        #region Fields

        private bool traceEnabled = true;

        #endregion

        #region Constructors

        public EwsTracer()
        {
            new EwsTracer(true);
        }

        public EwsTracer(bool enabled)
        {
            this.traceEnabled = enabled;
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Method inherited from ITraceListener. Write trace stream.
        /// </summary>
        /// <param name="traceType">Type of trace.</param>
        /// <param name="traceMessage">Trace message.</param>
        public void Trace(string traceType, string traceMessage)
        {
            if (this.traceEnabled)
                this.WriteTraceMessage(traceMessage, this.GetLoggingPath(traceType));
        }
                
        #endregion

        #region Private Methods

        /// <summary>
        /// Returns logging path for current traceType
        /// </summary>
        /// <param name="traceType">Trace type provided by Trace method.</param>
        /// <returns>string</returns>
        private string GetLoggingPath(string traceType)
        {
            string ewsFolder = "XEws";
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            string ewsPath = Path.Combine(desktopPath, ewsFolder);
            if (!Directory.Exists(ewsPath))
                Directory.CreateDirectory(ewsPath);

            return (string.Format("{0}",
                Path.Combine(ewsPath, (string.Format("{0} {1}.txt", traceType, DateTime.Now.ToString("hh-mm-ss-ms"))))
            ));
        }

        /// <summary>
        /// Write content of trace message to provided path.
        /// </summary>
        /// <param name="traceMessage">Message to write.</param>
        /// <param name="path">Path to save message to.</param>
        private void WriteTraceMessage(string traceMessage, string path)
        {
            File.WriteAllText(path, traceMessage);
        }

        #endregion

        #region Override Methods

        #endregion
    }
}
