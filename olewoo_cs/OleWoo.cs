using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices.ComTypes;

namespace olewoo_cs
{
    public partial class OleWoo : Form
    {
        private MRUList _mrufiles;

        public OleWoo()
        {
            InitializeComponent();
            tcTypeLibs.ImageList = imgListMisc;
            _mrufiles = new MRUList(@"Software\benf.org\olewoo\MRU");
            List<String> args = Environment.GetCommandLineArgs().ToList();
            args.RemoveAt(0);
            foreach (String arg in args)
            {
                OpenFile(System.IO.Path.GetFullPath(arg));
            }
        }

        private void OpenFile(String fname)
        {
            try
            {
                OWTypeLib tl = new OWTypeLib(fname);
                TabPage tp = new TabPage(tl.ShortName);
                tp.ImageIndex = 0;
                wooctrl wc = new wooctrl(imglstTreeNodes, imgListMisc, tl);
                tp.Controls.Add(wc);
                tp.Tag = tl;
                wc.Dock = System.Windows.Forms.DockStyle.Fill;
                tcTypeLibs.TabPages.Add(tp);
                _mrufiles.AddItem(fname);
                _mrufiles.Flush();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error:", MessageBoxButtons.OK);
            }
        }
        
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Dll files (*.dll)|*.dll|Typelibraries (*.tlb)|*.tlb|Executables (*.exe)|*.exe|ActiveX controls (*.ocx)|*.ocx|All files (*.*)|*.*";
            ofd.CheckFileExists = true;
            switch (ofd.ShowDialog(this))
            {
                case DialogResult.OK:
                    OpenFile(ofd.FileName);
                    break;
                default:
                    break;
            }
        }

        private void aboutOleWooToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox ab = new AboutBox();
            ab.ShowDialog();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void openMRUItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem mni = sender as ToolStripMenuItem;
            if (mni == null) return;
            OpenFile(mni.Tag as String);            
        }

        private void clearMRUItem_Click(object sender, EventArgs e)
        {
            _mrufiles.Clear();
            _mrufiles.Flush();
        }

        delegate void VoidDelg();

        private void fileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            /* Dynamically populate the file item with the MRU */
            this.fileToolStripMenuItem.DropDownItems.Clear();

            List<ToolStripItem> tsis = new List<ToolStripItem>();

            // 
            // openToolStripMenuItem
            // 
            ToolStripMenuItem tmiOpen = new ToolStripMenuItem();
            tmiOpen.Name = "openToolStripMenuItem";
            tmiOpen.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.O)));
            tmiOpen.Size = new System.Drawing.Size(208, 22);
            tmiOpen.Text = "&Open Typelibrary";
            tmiOpen.Click += new System.EventHandler(this.openToolStripMenuItem_Click);
            tsis.Add(tmiOpen);

            VoidDelg addSep = () =>
            {
                ToolStripSeparator tmiSep = new ToolStripSeparator();
                tmiSep.Name = "toolStripMenuItem1";
                tmiSep.Size = new System.Drawing.Size(205, 6);
                tsis.Add(tmiSep);
            };

            addSep();

            String[] mru = _mrufiles.Items;
            if (mru.Length > 0)
            {
                int idx = 1;
                foreach (String mrui in mru)
                {
                    ToolStripMenuItem tmiMru = new ToolStripMenuItem();
                    tmiMru.Tag = mrui;
                    tmiMru.Size = new System.Drawing.Size(208, 22);
                    String label = mrui;
                    if (label.Length > 35) label = label.Substring(0,10) + "..."+ label.Substring(label.Length - 20);
                    tmiMru.Text = "&" + (idx++) + " " + label;
                    tmiMru.Click += new System.EventHandler(this.openMRUItem_Click);
                    tsis.Add(tmiMru);
                }
                addSep();
                {
                    ToolStripMenuItem tmiMru = new ToolStripMenuItem();
                    tmiMru.Size = new System.Drawing.Size(208, 22);
                    tmiMru.Text = "&Clear Recent items list.";
                    tmiMru.Click += new System.EventHandler(this.clearMRUItem_Click);
                    tsis.Add(tmiMru);
                }
                addSep();
            }

            ToolStripMenuItem tmiExit = new ToolStripMenuItem();
            tmiExit.Name = "exitToolStripMenuItem";
            tmiExit.Size = new System.Drawing.Size(208, 22);
            tmiExit.Text = "E&xit";
            tmiExit.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            tsis.Add(tmiExit);

            this.fileToolStripMenuItem.DropDownItems.AddRange(tsis.ToArray());
        }

        private void registerContextHandlerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string menuCommand = string.Format("\"{0}\" \"%L\"",
                                               Application.ExecutablePath);
            FileShellExtension.Register("dllfile", "OleWoo Context Menu",
                                        "Open with OleWoo", menuCommand);
        }

        private void unregisterContextHandlerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileShellExtension.Unregister("dllfile", "OleWoo Context Menu");
        }

        private void tcTypeLibs_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
