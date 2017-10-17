/**************************************
 *
 * Part of OLEWOO - http://www.benf.org
 *
 * CopyLeft, but please credit.
 *
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace oledump
{
    class PlainIDLFormatter : olewoo_cs.IDLFormatter
    {
        StringBuilder _sb;
        bool _bPendingApplyTabs;

        public PlainIDLFormatter()
        {
            _sb = new StringBuilder();
            _bPendingApplyTabs = false;
        }

        public override void NewLine()
        {
            _sb.AppendLine("");
            _bPendingApplyTabs = true;
        }

        private void ApplyTabs()
        {
            _bPendingApplyTabs = false;
            if (_tabdepth > 0)
            {
                String s = "";
                for (int x = 0; x < _tabdepth; ++x)
                {
                    s += "\t";
                }
                _sb.Append(s);
            }
        }

        public override void AddString(String s)
        {
            if (_bPendingApplyTabs) ApplyTabs();
            _sb.Append(s);
        }

        public override void AddLink(String s, String o)
        {
            AddString(s);
        }

        public override string ToString()
        {
            return _sb.ToString();
        }
    }
    
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                if (args.Length < 1) throw new Exception("oledump TLBNAME");
                olewoo_cs.OWTypeLib tl = new olewoo_cs.OWTypeLib(args[0]);
                PlainIDLFormatter pi = new PlainIDLFormatter();
                tl.BuildIDLInto(pi);
                System.Console.WriteLine(pi.ToString());
            }
            catch (Exception e)
            {
                System.Console.WriteLine("OleDump:\r\n");
                System.Console.Error.WriteLine("Error : " + e.Message);
            }
        }
    }
}
