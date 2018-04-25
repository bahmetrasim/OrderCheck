using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using OfficeOpenXml;
using System.IO;
using System.Globalization;
using System.Xml;
using System.Text.RegularExpressions;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Style;

namespace Siparis_Kontrol
{
    public partial class Form1 : Form
    {
        string sheet = "";
        string pathins = "";
        string newsheet = "";
        Pivot prepare = new Pivot();

        public Form1()
        {
            InitializeComponent();
            backlog.Checked = true;
        }

        private void Choose_Click_1(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog ChoosingExcelFile = new OpenFileDialog();
                ChoosingExcelFile.Filter = "Excel Files | *.xlsx; *.xls; *.xlsm; *.xml";
                if (ChoosingExcelFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    this.Path.Text = ChoosingExcelFile.FileName;
                    pathins = Path.Text;
                }
                string connection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Path.Text + ";Extended Properties = \"Excel 12.0; HDR = YES;\"; ";
                OleDbConnection con = new OleDbConnection(connection);
                con.Open();
                ChooseSheets.DataSource = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                ChooseSheets.DisplayMember = "TABLE_NAME";
                ChooseSheets.ValueMember = "TABLE_NAME";
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void epplus_Click_1(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(Path.Text))
            {
                MessageBox.Show("Lütfen Dosyayı Seçiniz");
                return;
            }
            string pathsout = pathins.Substring(0, pathins.LastIndexOf("."));
            string ext = pathins.Substring(pathins.LastIndexOf("."));
            if (ext == ".xls")
            {
                MessageBox.Show("Lütfen Excel Yeni Versiyonu ile kaydediniz");
            }
            if (!radioButton1.Checked && !radioButton2.Checked)
            {
                MessageBox.Show("Lütfen Hangi Raporu Düzenlemek İstediğinizi Seçiniz");
                return;
            }
            pathsout = pathsout + "_1" + ext;
            sheet = ChooseSheets.SelectedValue.ToString();
            sheet = sheet.Substring(0, sheet.Length - 1);
            newsheet = sheet + "_1";
            int orderthick = int.Parse(tborderthick.Text);
            int paintwidth = int.Parse(tbpaintwidth.Text);
            int quantity = int.Parse(tbquantity.Text);
            int status = int.Parse(tbstatus.Text);
            string statusname = tbstatusname.Text;
            if (radioButton1.Checked)
            {
                bool nodate = false;
                bool all = false;
                if (allorders.Checked)
                {
                    all = true;
                }
                if (handling.Checked)
                {
                    nodate = true;
                }
                double netproduct = prepare.Pivotcalcqlik(pathins, sheet, pathsout, all, nodate);
                label3.Text = netproduct.ToString();
            }
            else
            {
                double netproduct = prepare.Pivotcalcaubt(pathins, sheet, newsheet, orderthick, paintwidth, quantity, status, statusname);
                label3.Text = netproduct.ToString();
            }
            MessageBox.Show("İşlem Tamamlanmıştır");
        }
    }
}
