
using MetroSet_UI.Forms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReConciler
{
    public partial class Main : MetroSetForm
    {
        frmLoadStatement statement; frmConvertStatement convertStatement; frmMergeSheets mergeSheets;
        public Main()
        {
            InitializeComponent();

            //Load version
            tssVersion.Text = "Version | " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
            tssVersion.ForeColor = Color.Maroon;
        }

        private void Statement_FormClosed(object sender, FormClosedEventArgs e)
        {
            statement = null;
    
        }

        private void toolStripBtnLoadStmt_Click(object sender, EventArgs e)
        {
            statement = new frmLoadStatement();
            statement.FormClosed += Statement_FormClosed;
            statement.MdiParent = this;
            statement.Show();
        }

        private void Main_Load(object sender, EventArgs e)
        {
            //MdiClient mdi;
            //foreach(Control ctl in this.Controls)
            //{
            //    mdi = (MdiClient)ctl;
            //    mdi.BackColor = System.Drawing.Color.WhiteSmoke;
            //}

            statement = new frmLoadStatement();
            statement.FormClosed += Statement_FormClosed;
            statement.MdiParent = this;
            statement.Show();
        }

        private void toolStripBtnConvert_Click(object sender, EventArgs e)
        {
            convertStatement = new frmConvertStatement();
            convertStatement.FormClosed += Statement_FormClosed;
            convertStatement.MdiParent = this;
            convertStatement.Show();
        }



        private void toolStripBtnMerge_Click(object sender, EventArgs e)
        {
            mergeSheets = new frmMergeSheets();
            mergeSheets.FormClosed += MergeSheets_FormClosed;
            mergeSheets.MdiParent = this;
            mergeSheets.Show();
        }

        private void MergeSheets_FormClosed(object sender, FormClosedEventArgs e)
        {
            mergeSheets = null;
        }
    }
}
