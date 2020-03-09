using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace etaxOnethVersion2.DATA_API
{
    class API_MAIL
    {
        public string email { get; set; }
        public string taxseller { get; set; }
        public string branch { get; set; }
        public string path { get; set; }
        public etaxOnethVersion2.etaxOneth form { get; set; }
        public string input { get; set; }
        public string timeuser { get; set; }
        public string typesoft { get; set; }
        public JObject jsontext { get; set; }
        public JObject jsontext_about { get; set; }
        public string numberstringgen { get; set; }
    }
}
