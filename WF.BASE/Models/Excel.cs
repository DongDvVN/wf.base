using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WF.BASE.Models
{
    public class Excel
    {
        public class Responce
        {
            public class ExportDataRequest
            {
                public string Id { get; set; }
                public string Title { get; set; }
                public string Content { get; set; }
            }
        }
    }
}
