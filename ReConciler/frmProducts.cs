using ReConciler.Data;
using ReConciler.Model;
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
    public partial class frmProducts : Form
    {
        public frmProducts()
        {
            InitializeComponent();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"D:\",
                Title = "Browse Text Files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "xls",
                Filter = "xls files (*.xlsx)|*.xlsx",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            var prodLst = DataAccess.GetProductList(textBox1.Text.Trim());
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = prodLst;
            dataGridView1.AutoSizeColumnsMode =
                DataGridViewAutoSizeColumnsMode.Fill;

            dataGridView1.AutoResizeColumns();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var prodLst = dataGridView1.DataSource as List<product>;


            prodLst.ForEach(p=>
            {
                ///Update tabe
                ///var getfromdb= context.Products.Where(p=>p.sku == p.productId).FistOrDefault();
                ///getfromdb.bv = p.bv;
                ///getfromdb.pv = p.pv;
                ///context.Entry(getfromdb).State = EntityState.Modified;


            });
            ///context.SaveChanges();
        }
    }
}
