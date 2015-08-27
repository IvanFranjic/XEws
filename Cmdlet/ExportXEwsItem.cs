using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using Microsoft.Exchange.WebServices.Data;
using System.Collections;
using XEws.CmdletAbstract;

namespace XEws.Cmdlet
{
    [Cmdlet(VerbsData.Export, "XEwsItem")]
    public class ExportXEwsItem : XEwsExportCmdlet
    {
        protected override void ProcessRecord()
        {
            this.ExportItem();
        }
    }

}