using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WF.BASE.Models
{
    public class Base
    {
        public class Request
        {
            public class GetAllRequest
            {
                public string Token { get; set; }
                public int Page { get; set; }
            }
        }
        public class Responce
        {
            public class GetAllRequest
            {
                public int code { get; set; }
                public string message { get; set; }
                public object data { get; set; }
                public long page { get; set; }
                public long items_per_page { get; set; }
                public string total_items { get; set; }
                public List<Request> requests { get; set; }
            }
            public class Request
            {
                public string id { get; set; }
                public string hid { get; set; }
                public long gid { get; set; }
                public string name { get; set; }
                public string content { get; set; }
                public List<Form> form { get; set; }
                public List<File> files { get; set; }
            }
            public class Form
            {
                public string id { get; set; }
                public string name { get; set; }
                public string value { get; set; }
                public string type { get; set; }
                public string placeholder { get; set; }
            }
            public class File
            {
                public string id { get; set; }
                public string name { get; set; }
                public long image { get; set; }
                public long size { get; set; }
                public long since { get; set; }
                public string username { get; set; }
                public string abs_url { get; set; }
                public string hid { get; set; }
                public string url { get; set; }
                public string download { get; set; }
            }
        }
    }
}
