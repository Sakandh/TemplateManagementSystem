using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace TMS_Framework.Models
{
    public class UnitOwner
    {
        public string barCode { get; set; }
        public string baseTemplateFileName { get; set; }

        public string mergedFileName { get; set; }

        public List<ClientData> clientDataList { get; set; }

        public List<TableData> tableDataList { get; set; }

        public List<BlockData> blockDataList { get; set; }
    }
}