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
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace olewoo_cs
{
    public partial class FindDialog : Form
    {
        PnlOleText _textctrl;
        public FindDialog(PnlOleText textctrl)
        {
            InitializeComponent();
            _textctrl = textctrl;
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnFindNext_Click(object sender, EventArgs e)
        {
            if (_textctrl.FindNextText(txtSearchString.Text, rbFindDown.Checked))
            {
//                this.Close();
            }
            else
            {
                MessageBox.Show("Cannot find '" + txtSearchString.Text + "'", "OleWoo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
