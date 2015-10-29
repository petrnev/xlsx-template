using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Windows.Forms.VisualStyles;
using System.Collections.Generic;
using System.Configuration;
using Newtonsoft.Json;
using System.Web;
using System.Data.OleDb;
using System.Reflection;
using ClosedXML.Excel;
using Microsoft.Office.Interop.Excel;

namespace xlxs_template
{
    public partial class Form1 : Form
    {

        public System.Data.DataTable dt;
        
        private string Excel03ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        private string Excel07ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        public Form1()
        {
            InitializeComponent();
        }

         private void label2_Click(object sender, EventArgs e)
        {}
         private void panel1_Paint(object sender, EventArgs e)
         { }

        private void Form1_Load(object sender, EventArgs e)
        {
           List<Item> items= LoadJson();
           comboBox1.Items.Add("Select Company");
            for (int i = 0; i < items.Count; i++)
           {
               comboBox1.Items.Add(items[i].companyname);
           }
            comboBox1.SelectedIndex = 0;
            exportbtn.Enabled = false;
        }

        public void populateCombo(List<Item> items)
        {
            comboBox1.Items.Clear();
            comboBox1.Items.Add("Select Company");
            for (int i = 0; i < items.Count; i++)
            {
                comboBox1.Items.Add(items[i].companyname);
            }
            comboBox1.SelectedIndex = 0;
        }

        private void Import_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            openFileDialog1.Filter = "allfiles|*.xlsx";
            if (result == DialogResult.OK)
            {
                lblError.Text = "File Uploaded";
                lblError.ForeColor = System.Drawing.Color.Green;
            }
            //int count = 0;
            string name = openFileDialog1.FileName;
            string ext = System.IO.Path.GetExtension(openFileDialog1.FileName).ToLower();   
            if (!ext.Equals(".xlsx"))
            {
                lblError.Text = "File format should be xls or xlsx";
                lblError.ForeColor = System.Drawing.Color.Red;
                return;
            }

            string conStr, sheetName;
            string header =  "NO";
            conStr = string.Empty;
            switch (ext)
            {

                case ".xls": //Excel 97-03
                    conStr = string.Format(Excel03ConString, name, header);
                    break;

                case ".xlsx": //Excel 07
                    conStr = string.Format(Excel07ConString, name, header);
                    break;
            }

            using (OleDbConnection con = new OleDbConnection(conStr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    cmd.Connection = con;
                    con.Open();
                    System.Data.DataTable dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                    con.Close();
                }
            }
           // DataTable dt = new DataTable();
            dt = new System.Data.DataTable();
            using (OleDbConnection con = new OleDbConnection(conStr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    using (OleDbDataAdapter oda = new OleDbDataAdapter())
                    {
                        
                        cmd.CommandText = "SELECT * From [" + sheetName + "]";
                        cmd.Connection = con;
                        con.Open();
                        oda.SelectCommand = cmd;
                        oda.Fill(dt);
                        con.Close();
                        exportbtn.Enabled = true;
                        //Populate DataGridView.
                        //dataGridView1.DataSource = dt;
                    }
                }
            }



           
            ////openFileDialog1.FileName.
            //foreach (string item in openFileDialog1.FileNames)
            //{
            //    FilenameName = item.Split('\\');

            //    //File.Copy(item, @"Images\" + FilenameName[FilenameName.Length - 1]);
            //    count++;
            //}
            //MessageBox.Show(Convert.ToString(count) + " File(s) copied");
        }

        public List<Item> LoadJson()
        {

            using (StreamReader r = new StreamReader(@"D:\companies.json"))
            {
                string json = r.ReadToEnd();
                
                
                List<Item> items = JsonConvert.DeserializeObject<List<Item>>(json);
                return items;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //List<Item> items = LoadJson();
            if (comboBox1.SelectedIndex != 0)
            {

                importbtn.Enabled = true;
                List<Item> items = LoadJson();
                string name=comboBox1.SelectedItem.ToString();
                for (int i = 0; i < items.Count; i++)
                {
                    if (name == items[i].companyname)
                    {
                        cmpname.Text = items[i].companyname;
                        city.Text = items[i].city;
                        street.Text = items[i].street;
                        description.Text = items[i].description;
                        region.Text = items[i].region;
                    }
                }
            }
            else 
            {
                    
                        cmpname.Text ="" ;
                        city.Text ="" ;
                        street.Text = "";
                        description.Text = "";
                        region.Text = "";
                        importbtn.Enabled = false;
            }

        }

        private void savebtn_Click(object sender, EventArgs e)
        {
            List<Item> items = LoadJson();
           
            int c=items.Count;
            c=c+1;
            if (comboBox1.SelectedIndex != 0)
            {
                for (int i = 0; i < items.Count; i++)
                {
                    if (items[i].companyname == cmpname.Text)
                    {
                        c= Int32.Parse(items[i].id);
                        items.Remove(items[i]);

                        items.Add(new Item()
                        {
                            companyname = cmpname.Text,
                            description = description.Text,
                            street = street.Text,
                            region = region.Text,
                            id = c.ToString(),
                            city = city.Text

                        });
                    }
                }
            }
            else 
            {
                items.Add(new Item()
                {
                    companyname = cmpname.Text,
                    description = description.Text,
                    street = street.Text,
                    region = region.Text,
                    id = c.ToString(),
                    city = city.Text

                });
                    
            }
            string json = JsonConvert.SerializeObject(items.ToArray());

            //write string to file
            System.IO.File.WriteAllText(@"D:\companies.json", json);

            List<Item> newitems= LoadJson();
            populateCombo(newitems);
        }

        private void deletbtn_Click(object sender, EventArgs e)
        {
            string cmp_name= comboBox1.SelectedItem.ToString();
            List<Item> items = LoadJson();
            for (int i = 0; i < items.Count; i++)
            {
                if (items[i].companyname == cmp_name)
                {
                    items.Remove(items[i]);
                }
            }
            string json = JsonConvert.SerializeObject(items.ToArray());

            //write string to file
            System.IO.File.WriteAllText(@"D:\companies.json", json);

            List<Item> newitems=LoadJson();
            populateCombo(newitems);
        }

        private void exportbtn_Click(object sender, EventArgs e)
        {
           
            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(missing);

            Microsoft.Office.Interop.Excel.Worksheet newWorksheet;
            newWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)excel.Worksheets.Add(missing, missing, missing, missing);
            newWorksheet.Name = "Name of data sheet";

            //  for first datatable dt1..


            System.Data.DataTable info = new System.Data.DataTable();
            info.Columns.AddRange(new DataColumn[4] { 
            new DataColumn("Hours", typeof(string)),
            new DataColumn("person name",typeof(string)),
            new DataColumn("person second name",typeof(string)),
            new DataColumn("division",typeof(string))
            
            });

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                info.Rows.Add(dt.Rows[i]["F6"], dt.Rows[i]["F1"], dt.Rows[i]["F2"], dt.Rows[i]["F3"]);
            }


           

            int iRow1 = 4;
            foreach (DataRow r in info.Rows)
            {
                iRow1++;

                for (int i = 1; i < info.Columns.Count + 1; i++)
                {
                    if (iRow1 <=8)
                    {
                    }
                    else
                    {
                        excel.Cells[iRow1 + 1, i + 1] = r[i - 1].ToString();
                    }
                }

            }

            
            int iCol1 = 1;
            foreach (DataColumn c in info.Columns)
            {
                
                iCol1++;
                excel.Cells[9, iCol1] = c.ColumnName;
                excel.Cells.Font.Bold = true;
            }

            excel.Cells.Font.Bold = false;


            //  for  second datatable dt2..
            int iCol2 = 0;
            System.Data.DataTable expdt = new System.Data.DataTable();
            expdt.Columns.AddRange(new DataColumn[2] { 
            new DataColumn("ColumnName", typeof(string)),
            new DataColumn("Column Value",typeof(string)) });
            expdt.Rows.Add("Company Name:", cmpname.Text);
            expdt.Rows.Add("Description:", description.Text);
            expdt.Rows.Add("City:", city.Text);
            expdt.Rows.Add("Region:", region.Text);
            expdt.Rows.Add("Street:", street.Text);

            //foreach (DataColumn c in expdt.Columns)
            //{
            //    iCol2++;
            //    excel.Cells[1, iCol2] = c.ColumnName;
            //}


            int iRow2 = -1;
            foreach (DataRow r in expdt.Rows)
            {
                iRow2++;

                for (int i = 1; i < expdt.Columns.Count + 1; i++)
                {

                    //if (iRow2 == 1)
                    //{
                    //    // Add the header the first time through 
                    //    //excel.Cells[iRow2, i] = expdt.Columns[i - 1].ColumnName;
                    //}
                    
                    excel.Cells[iRow2 + 1, i] = r[i - 1].ToString();
                    
                }

            }
            //Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal
            workbook.SaveAs("C:\\Excel\\exampleexport"+ DateTime.Now.Second+".xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, missing, missing, missing, missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, missing, missing, missing, missing, missing);
            workbook.Close(true, missing, missing);
            excel.Quit();
            lblError.Text = "File Exported Successfully";
            lblError.ForeColor = System.Drawing.Color.Green;
            #region export old code
            ////foreach (DataGridViewRow row in dataGridView1.Rows)
            ////{
            ////    dt.Rows.Add();
            ////    foreach (DataGridViewCell cell in row.Cells)
            ////    {
            ////        dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex] = cell.Value.ToString();
            ////    }
            ////}
            //System.Data.DataTable expdt = new System.Data.DataTable();
            //expdt.Columns.AddRange(new DataColumn[3] { new DataColumn("Id", typeof(int)),
            //new DataColumn("Name", typeof(string)),
            //new DataColumn("Country",typeof(string)) });
            //expdt.Rows.Add(1, "John Hammond", "United States");
            //expdt.Rows.Add(2, "Mudassar Khan", "India");
            //expdt.Rows.Add(3, "Suzanne Mathews", "France");
            //expdt.Rows.Add(4, "Robert Schidner", "Russia");
            ////Exporting to Excel
            //DataSet ds = new DataSet();
            //ds.Tables.Add(dt);
            //ds.Tables.Add(expdt);

            //string folderPath = "C:\\Excel\\";
            //if (!Directory.Exists(folderPath))
            //{
            //    Directory.CreateDirectory(folderPath);
            //}
            //using (XLWorkbook wb = new XLWorkbook())
            //{
            //    wb.Worksheets.Add(ds);
            //    //wb.Cells("abc"). = "abc";
            //    //wb.Worksheets.Add(dt, "Customers");


            //    wb.SaveAs(folderPath + "DataGridViewExport.xlsx");
            //} 
            #endregion

            exportbtn.Enabled = false;
        }

        
    }
    public class Item
    {
        public string companyname;
        public string description;
        public string city;
        public string street;
        public string region;
        public string id;
        
    }



}



