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

namespace Excell_Reader
{
    public partial class Form1 : Form
    {


        public Form1()
        {
            InitializeComponent();

        }

      public void button1_Click(object sender, EventArgs e)
        {

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Excell sec";
            ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            ofd.Filter = "Excell(*xls.xlsx)|*.xls;*.xlsx";
            ofd.ShowDialog();

            string link = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;" +
                "Data Source={0};Extended Properties=Excel 12.0 Xml; ", ofd.FileName);
            OleDbConnection con = new OleDbConnection(link);
            label1.Text = ofd.FileName;
            con.Open();
           DataTable dt = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            con.Close();
            comboBox1.Items.Clear();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string sheetname = dt.Rows[i]["TABLE_NAME"].ToString();
                sheetname = sheetname.Substring(1, sheetname.Length - 3);
                sheetname = sheetname.Replace("#", "");
                comboBox1.Items.Add(sheetname);
                con.Close();
            }
            label1.Visible = false;

        }

      private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
              


                string link = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;" +
                  "Data Source={0};Extended Properties=Excel 12.0 Xml; ", label1.Text) ;
                OleDbConnection elaqe = new OleDbConnection(link);
               

                if (comboBox1.Text == "")
                {
                    MessageBox.Show("Zəhmət olmasa ərazi seçin.", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                if (comboBox1.Text != "")
                {
                    OleDbDataAdapter sehife = new OleDbDataAdapter("Select * from [" + comboBox1.Text + "$]", elaqe);
                    DataTable cedvel = new DataTable();
                    sehife.Fill(cedvel);
                    dataGridView1.DataSource = cedvel;
                    elaqe.Close();

                }
            }
            catch (Exception)
            {

                MessageBox.Show("Təəssüf ki bazada məlumat yoxdur");
            }

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult cavab = MessageBox.Show("Are you sure?", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (cavab==DialogResult.No)
            {
                e.Cancel = true;
                   
            }
        }

        
        private void button2_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() != DialogResult.Cancel) 
            {
                Form1 frm1 = new Form1();
               
                button1.BackColor = colorDialog1.Color;
                button2.BackColor = colorDialog1.Color;
                panel1.BackColor = colorDialog1.Color;
                panel2.BackColor = colorDialog1.Color;
                dataGridView1.BackgroundColor = colorDialog1.Color;
                dataGridView1.GridColor = colorDialog1.Color;
                dataGridView1.ForeColor = colorDialog1.Color;
          

                
            }
        }
    }
}
