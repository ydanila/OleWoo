﻿/**************************************
 *
 * Part of OLEWOO - http://www.benf.org
 *
 * CopyLeft, but please credit.
 *
 */
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace olewoo_cs
{
    public partial class wooctrl : UserControl
    {
        enum SortType
        {
            Sorted_Numerically,
            Sorted_AlphaUp,
            Sorted_AlphaDown,
            Sorted_Max
        }

        OWTypeLib _tlib;
        NodeLocator _nl;
        ImageList _iml;
        SortType _sort;

        public wooctrl(ImageList imglstTreeNodes, ImageList imglstMisc, OWTypeLib tlib)
        {
            _tlib = tlib;
            _nl = new NodeLocator();
            _iml = imglstMisc;
            _sort = SortType.Sorted_Numerically;

            InitializeComponent();
            this.txtOleDescrPlain.ParentCtrl = this;
            tvLibDisp.ImageList = imglstTreeNodes;
            this.Dock = DockStyle.Fill;

            tvLibDisp.Nodes.Add(GenNodeTree(tlib, _nl));
            this.txtOleDescrPlain.NodeLocator = _nl;
            tvLibDisp.Nodes[0].Expand();
        }

        public ImageList ImageList
        {
            get
            {
                return _iml;
            }
        }

        /*
         * Note that this generates redundant tree nodes, i.e. many definitions of (eg) IUnknown
         * this is because Trees don't have support for child sharing. (how would the parent property work? :)
         */
        private TreeNode GenNodeTree(ITlibNode tln, NodeLocator nl)
        {
            TreeNode tn = new TreeNode(tln.Name, tln.ImageIndex, 
                (int)ITlibNode.ImageIndices.idx_selected, 
                tln.Children.ConvertAll(x => GenNodeTree(x,nl)).ToArray());
            tn.Tag = tln;
            nl.Add(tn);
            return tn;
        }

        private void tvLibDisp_AfterSelect(object sender, TreeViewEventArgs e)
        {
            txtOleDescrPlain.SetCurrentNode(e.Node);
        }

        private void ClearMatches()
        {
            pnlMatchesList.Visible = false;
            lstNodeMatches.Items.Clear();
        }
        /*
         * Search through the registered names for the tree nodes.
         * 
         * When we hit one, select that node.
         */
        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            String text = txtSearch.Text;
            if (text == "")
            {
                ClearMatches();
            }
            else
            {
                try
                {
                    List<NamedNode> tn = _nl.FindMatches(text);
                    lstNodeMatches.Items.Clear();
                    if (tn != null && tn.Count > 0)
                    {
                        tvLibDisp.ActivateNode(tn.First().TreeNode);
                        lstNodeMatches.Items.AddRange(tn.ToArray());
                    }
                    pnlMatchesList.Visible = true;
                }
                catch (Exception)
                {
                }
            }
        }

        private void lstNodeMatches_SelectedIndexChanged(object sender, EventArgs e)
        {
            NamedNode nn = lstNodeMatches.SelectedItem as NamedNode;
            if (nn == null) return;
            tvLibDisp.ActivateNode(nn.TreeNode);
        }

        private void btnHideMatches_Click(object sender, EventArgs e)
        {
            txtSearch.Text = "";
        }

        private void btnAddNodeTabl_Click(object sender, EventArgs e)
        {
            txtOleDescrPlain.AddTab(tvLibDisp.SelectedNode);
        }
        public void SelectTreeNode(TreeNode tn)
        {
            tvLibDisp.ActivateNode(tn);
        }

        delegate int CompDelg(ITlibNode x, ITlibNode y);

        class OleNodeComparer : IComparer<TreeNode>
        {
            CompDelg _cd;
            public OleNodeComparer(SortType t)
            {
                switch (t)
                {
                    case SortType.Sorted_Numerically:
                    default: // shouldn't happen...
                        _cd = CompareNum;
                        break;
                    case SortType.Sorted_AlphaUp:
                        _cd = CompareAlphaUp;
                        break;
                    case SortType.Sorted_AlphaDown:
                        _cd = CompareAlphaDown;
                        break;
                }
            }
            public int Compare(TreeNode x, TreeNode y)
            {
                return _cd(x.Tag as ITlibNode, y.Tag as ITlibNode);
            }
            int CompareNum(ITlibNode x, ITlibNode y)
            {
                return x.Idx.CompareTo(y.Idx);
            }
            int CompareAlphaUp(ITlibNode x, ITlibNode y)
            {
                return x.Name.CompareTo(y.Name);
            }
            int CompareAlphaDown(ITlibNode x, ITlibNode y)
            {
                return y.Name.CompareTo(x.Name);
            }
        }

        private void btnSortAlpha_Click(object sender, EventArgs e)
        {
            using (new UpdateSuspender(tvLibDisp))
            {
                _sort++;
                if (_sort == SortType.Sorted_Max) _sort = SortType.Sorted_Numerically;
                List<TreeNode> nlst = new List<TreeNode>();
                TreeNode root = tvLibDisp.Nodes[0];
                System.Collections.IEnumerator enm = root.Nodes.GetEnumerator();
                while (enm.MoveNext())
                {
                    nlst.Add(enm.Current as TreeNode);
                }
                nlst.Sort(new OleNodeComparer(_sort));
                root.Nodes.Clear();
                root.Nodes.AddRange(nlst.ToArray());
                //  update children order
                var r = (ITlibNode)root.Tag;
                r.Children = (from n in nlst select (ITlibNode)n.Tag).ToList();
            }
        }

        private void btnRewind_Click(object sender, EventArgs e)
        {
            txtOleDescrPlain.RewindOne();
        }

    }

    static class xtras
    {
        public static void ActivateNode(this TreeView tv, TreeNode tn)
        {
            if (tn != null)
            {
                tv.SelectedNode = tn;
                tn.EnsureVisible();
            }
        }
    }

    public class NamedNode
    {
        private TreeNode _tn;
        private ITlibNode _tln;
        public NamedNode(TreeNode tn)
        {
            _tn = tn;
            _tln = tn.Tag as ITlibNode;
        }
        public String Name
        {
            get
            {
                return _tln.ShortName;
            }
        }
        public String ObjectName
        {
            get
            {
                return _tln.ObjectName;
            }
        }
        public TreeNode TreeNode
        {
            get
            {
                return _tn;
            }
        }
        public override string ToString()
        {
            return _tln.ShortName;
        }
    }

    public class NodeLocator
    {
        List<NamedNode> nodes;
        Dictionary<String, NamedNode> linkmap;

        public NodeLocator()
        {
            nodes = new List<NamedNode>();
            linkmap = new Dictionary<string, NamedNode>();
        }

        public void Add(TreeNode tn)
        {
            ITlibNode tli = tn.Tag as ITlibNode;
            String name = tli.ShortName;
            if (name != null)
            {               
                NamedNode nn = new NamedNode(tn);
                nodes.Add(nn);
                String oname = tli.ObjectName;
                if (!String.IsNullOrEmpty(oname) && !linkmap.ContainsKey(oname))
                {
                    linkmap[oname] = nn;
                }
            }
        }
        // O(N).  FIX!
        public List<NamedNode> FindMatches(String text)
        {
            Regex re = new Regex("^.*" + text, RegexOptions.IgnoreCase);
            return nodes.FindAll(x => re.IsMatch(x.Name));
        }

        public NamedNode FindLinkMatch(String text)
        {
            if (linkmap.ContainsKey(text))
            {
                return linkmap[text];
            }
            return null;
        }
    }
}
