using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArrowGlobal
{
    public class PhoneNumber
    {
        private string mobile;
        private string work;
        private string home;

        public string ArrowKey { get; set; }

        public string Mobile
        {
            get { return mobile; }
            set { mobile = value != null ? value.PadLeft(11, '0') : string.Empty; }
        }

        public string Work
        {
            get { return work; }
            set { work = value != null ? value.PadLeft(11, '0') : string.Empty; }
        }

        public string Home
        {
            get { return home; }
            set { home = value != null ? value.PadLeft(11, '0') : string.Empty; }
        }
    }
}
