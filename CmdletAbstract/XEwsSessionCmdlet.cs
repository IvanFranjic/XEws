using System;
using System.Management.Automation;
using Microsoft.Exchange.WebServices.Data;
using System.IO;
using System.Security;

namespace XEws.CmdletAbstract
{
    public abstract class XEwsSessionCmdlet : XEwsCmdlet
    {
        private string autodiscoverEmail = null;
        [Parameter(Position = 0, ParameterSetName = "Autodiscover")]
        public string AutodiscoverEmail
        {
            get
            {
                return autodiscoverEmail;
            }
            set
            {
                autodiscoverEmail = value;
            }
        }

        private Uri ewsUri = null;
        [Parameter(Position = 0, ParameterSetName = "ManualUrl")]
        public Uri EwsUri
        {
            get
            {
                return ewsUri;
            }
            set
            {
                ewsUri = value;
            }
        }

        [Parameter(Position = 1, Mandatory = true)]
        public string UserName
        {
            get;
            set;
        }

        [Parameter(Position = 2, Mandatory = true)]
        public SecureString Password
        {
            get;
            set;
        }
        
        [Parameter(Position = 3, Mandatory = false)]
        public string ImpersonationEmail
        {
            get;
            set;
        }

        internal bool traceEnabled = false;
        [Parameter(Position = 4, Mandatory = false)]
        public bool TraceEnabled
        {
            get
            {
                return traceEnabled;
            }
            set
            {
                traceEnabled = value;
            }
        }

        internal string traceOutputFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        [Parameter(Position = 5, Mandatory = false)]
        public string TraceOutputFolder
        {
            get
            {
                return traceOutputFolder;
            }
            set
            {
                traceOutputFolder = value;
            }
        }

        internal TraceFlags traceFlags = TraceFlags.EwsResponse;
        [Parameter(Position = 6, Mandatory = false)]
        public TraceFlags TraceFlags
        {
            get
            {
                return traceFlags;
            }
            set
            {
                traceFlags = value;
            }
        }

        /// <summary>
        /// Method to check if trace path exist.
        /// </summary>
        /// <param name="tracePath">Path of the folder where traces should be placed.</param>
        internal void ValidateTracePath(string tracePath)
        {
            if (Directory.Exists(tracePath))
                return;

            throw new DirectoryNotFoundException(String.Format("Directory '{0}' is not valid path."));
        }
    }
}
