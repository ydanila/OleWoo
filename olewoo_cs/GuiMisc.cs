using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace olewoo_cs
{
    delegate void EndUpdateDelg();

    class UpdateSuspender : IDisposable
    {
        EndUpdateDelg _eud;
        public UpdateSuspender(TreeView t)
        {
            t.BeginUpdate();
            _eud = t.EndUpdate;
        }
        public void Dispose()
        {
            _eud();
        }
    }
    class TBUpdateSuspender : IDisposable
    {
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private extern static IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wp, IntPtr lp); 
        
        TextBoxBase _tbb;
        public TBUpdateSuspender(TextBoxBase tbb)
        {
            _tbb = tbb;
            SendMessage(_tbb.Handle, 0xb, (IntPtr)0, IntPtr.Zero); 
        }
        public void Dispose()
        {
            SendMessage(_tbb.Handle, 0xb, (IntPtr)1, IntPtr.Zero);
            _tbb.Invalidate();
        }
    }
}
