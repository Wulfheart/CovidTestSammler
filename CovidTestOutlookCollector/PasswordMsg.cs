using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CovidTestOutlookCollector
{
    class PasswordMsg : Msg
    {
        public string Password { get
            {
                Regex rx = new Regex(@"(?<=\bDas Passwort lautet:\s)(.*)");
                string matched = rx.Match(Message.BodyText).ToString();
                return string.IsNullOrEmpty(matched) ? matched : new string(matched.Where(x => !char.IsWhiteSpace(x)).ToArray());
            }
        }

        public string PdfFileName
        {
            get
            {
                Regex rx = new Regex(@"(?<=\bHiermit enthalten Sie das Passwort zum Öffnen der PDF-Datei\s)(\w+)");
                return rx.Match(Message.BodyText).ToString() + ".pdf";
            }
        }
        public PasswordMsg(string path) : base(path)
        {
        }
    }
}
