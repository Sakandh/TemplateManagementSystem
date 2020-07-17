using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace TMS_Framework.Models
{
    public class TableData
    {
        public string Key { get; set; }
        public string Value { get; set; }

        public int rowCount { get; set; }

        public int columnCount { get; set; }
    }
}