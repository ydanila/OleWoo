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

namespace olewoo_cs
{
    public abstract class IDLFormatter : olewoo_interop.IDLFormatter_iop
    {
        protected int _tabdepth;
        public void Indent() { _tabdepth++; }
        public void Dedent() { _tabdepth--; }
        public abstract void NewLine();
        public void AppendLine(String s)
        {
            AddString(s);
            NewLine();
        }
    }

    class IDLHelperTab : IDisposable
    {
        IDLFormatter _ih;

        public IDLHelperTab(IDLFormatter ih)
        {
            ih.Indent();
            _ih = ih;
        }
        public void Dispose()
        {
            _ih.Dedent();
        }
    }
}
