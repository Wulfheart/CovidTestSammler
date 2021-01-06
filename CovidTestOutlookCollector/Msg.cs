using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CovidTestOutlookCollector
{
    class Msg
    {
        public string Path { get; set; }
        public string ID { get; set; }

        public MsgReader.Outlook.Storage.Message Message { get; set; }

        public Msg()
        {

        }
        public Msg(string path)
        {
            Path = path;
            Message = new MsgReader.Outlook.Storage.Message(Path);
            Regex rx = new Regex(@"(DE)\S*");
            ID = rx.Match(Message.Subject).Value;

            // TODO: Throw error if ID is null
        }
    }
}
