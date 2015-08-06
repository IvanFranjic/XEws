using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using Microsoft.Exchange.WebServices.Data;
using System.IO;

namespace XEws.CmdletAbstract
{
    public abstract class XEwsSessionCmdlet : XEwsCmdlet
    {
        [Parameter()]
        public string UserName
        {
            get;
            set;
        }

        [Parameter()]
        public string Password
        {
            get;
            set;
        }

        [Parameter()]
        public Uri EwsUri
        {
            get;
            set;
        }

        [Parameter()]
        public string ImpersonationEmail
        {
            get;
            set;
        }

        internal bool traceEnabled = false;
        [Parameter()]
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
        [Parameter()]
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
        [Parameter()]
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
