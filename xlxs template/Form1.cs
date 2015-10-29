﻿using System;
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
using System.Configuration;

namespace xlxs_template
{
    public partial class Form1 : Form
    {
        public List<Setting> _newset;
        public System.Data.DataTable dt;
        public System.Data.DataTable ImportDataTable;
        public string filen;
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
            
            exportbtn.Enabled = false;
           List<Item> items= LoadJson();
           comboBox1.Items.Add("Select Company");
            for (int i = 0; i < items.Count; i++)
           {
               comboBox1.Items.Add(items[i].companyname);
           }
            comboBox1.SelectedIndex = 0;
           
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
            Setting lst = LoadJsonSettings();
            DialogResult result = openFileDialog1.ShowDialog();
            openFileDialog1.Filter = "allfiles|*.xlsx";
            if (result == DialogResult.OK)
            {
                lblError.Text = "File Uploaded";
                lblError.ForeColor = System.Drawing.Color.Green;
            }

            if(!importMethod(lst))
            {
                return;
            }
            string name = openFileDialog1.FileName;
            string ext = System.IO.Path.GetExtension(openFileDialog1.FileName).ToLower();   
            if (!ext.Equals(".xlsx"))
            {
                lblError.Text = "File format should be xlsx";
                lblError.ForeColor = System.Drawing.Color.Red;
                return;
            }
            string[] _nameofFile = new string[5] ;
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
                    int SheetNumber = Int32.Parse(System.Configuration.ConfigurationManager.AppSettings["SheetNumber"]);
                    sheetName = dtExcelSchema.Rows[SheetNumber-1]["TABLE_NAME"].ToString();
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

                        string[] limit= lst.exportColumnsLocation.Split(':');
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



            #region -new code for import

            System.Data.DataTable _impdt = ImportDataTable;
            string[] valname = new string[dt.Rows.Count];
            for (int a = 0; a < dt.Rows.Count; a++)
            {
                for (int b = 0; b < dt.Columns.Count; b++)
                {

                    for (int f = 0; f < _impdt.Columns.Count; f++)
                    {

                        if (dt.Rows[a][b].ToString() == _impdt.Columns[f].ColumnName)
                        {

                            //int remainig = a;
                            //string[] valname=new  string[dt.Rows.Count-a];
                            int d = 0;
                            for (int c = a; c < dt.Rows.Count; c++)
                            {
                                if (valname[d] == null)
                                {
                                    valname[d] = dt.Rows[c][b].ToString();
                                }
                                else
                                {
                                    valname[d] = valname[d] + "," + dt.Rows[c][b].ToString();
                                }
                                d++;
                            }

                        }

                    }

                    #region -oldcode
                    //if (dt.Rows[a][b].ToString() == "name" || dt.Rows[a][b].ToString() == "surname" || dt.Rows[a][b].ToString() == "Divisions")
                    //{

                    //    int remainig = a;
                    //    //string[] valname=new  string[dt.Rows.Count-a];
                    //    int d = 0;
                    //    for (int c = a + 1; c < dt.Rows.Count; c++)
                    //    {
                    //        if (valname[d] == null)
                    //        {
                    //            valname[d] = dt.Rows[c][b].ToString();
                    //        }
                    //        else
                    //        {
                    //            valname[d] = valname[d] + "," + dt.Rows[c][b].ToString();
                    //        }
                    //        d++;
                    //    }

                    //}
                    #endregion

                }
            }

            Setting _setting = LoadJsonSettings();
            string[] newarray = valname.Where(c => c != null).ToArray();

            for (int h = 0; h < newarray.Count(); h++)
            {
                string[] val = newarray[h].Split(',');
                if (h == 0)
                {
                    _impdt = DataTableExtensions.SetColumnsOrder(_impdt, val);
                    //DataTableExtensions dte = new DataTableExtensions();
                    //_impdt = dte.SetColumnsOrder(_impdt, val);
                }
                DataRow row = _impdt.NewRow();

                for (int g = 0; g < _impdt.Columns.Count; g++)
                {
                    row[g] = val[g];
                }

                _impdt.Rows.Add(row);
            }


            #endregion

            ////openFileDialog1.FileName.
            foreach (string item in openFileDialog1.FileNames)
            {
                _nameofFile = item.Split('\\');

                //File.Copy(item, @"Images\" + FilenameName[FilenameName.Length - 1]);
                //count++;
            }

            string[] newname = _nameofFile[_nameofFile.Length - 1].Split('.');
            filen = newname[0];
            
          
        }

        public List<Item> LoadJson()
        {
            string path = System.Configuration.ConfigurationManager.AppSettings["companyFilePath"];
            if (!File.Exists(path))
            {
                lblError.Text = "companies.Json file does not exist";
                lblError.ForeColor = System.Drawing.Color.Red;
                return null;   
            }

            lblError.Text = string.Empty;
            using (StreamReader r = new StreamReader(path))
            {
                string json = r.ReadToEnd();
                
                
                List<Item> items = JsonConvert.DeserializeObject<List<Item>>(json);
                return items;
            }
        }

        public void LoadSettinJson()
        {
            string path = System.Configuration.ConfigurationManager.AppSettings["settingFilePath"];
            if (!File.Exists(path))
            {
                lblError.Text = "companies.Json file does not exist";
                lblError.ForeColor = System.Drawing.Color.Red;
                //return null;
            }
           
            lblError.Text = string.Empty;
            using (StreamReader rs = new StreamReader(path))
            {
                string jsons = rs.ReadToEnd();
                object target=null;
               // List<Setting> _settings = 
                    JsonConvert.PopulateObject(jsons, target);//.DeserializeObject<List<Setting>>(jsons);
                //_newset = _settings;
               // return _settings;
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
            if (items == null)
            {
                lblError.Text = "You cannot save data now";
                lblError.ForeColor = System.Drawing.Color.Red;
                return;
            }

            bool _isupdate = false;
            int c=items.Count;
            c=c+1;
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
                        _isupdate = true;
                    }
                    
                }
                if (!_isupdate)
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
            //System.IO.File.WriteAllText(@"D:\companies.json", json);
            System.IO.File.WriteAllText(System.Configuration.ConfigurationManager.AppSettings["companyFilePath"], json);

            List<Item> newitems= LoadJson();
            populateCombo(newitems);
        }

        private void deletbtn_Click(object sender, EventArgs e)
        {
            string cmp_name= comboBox1.SelectedItem.ToString();
            List<Item> items = LoadJson();
            if (items == null)
            {
                lblError.Text = "You cannot delete data now";
                lblError.ForeColor = System.Drawing.Color.Red;
                return;
            }
            for (int i = 0; i < items.Count; i++)
            {
                if (items[i].companyname == cmp_name)
                {
                    items.Remove(items[i]);
                }
            }
            string json = JsonConvert.SerializeObject(items.ToArray());

            //write string to file
            System.IO.File.WriteAllText(System.Configuration.ConfigurationManager.AppSettings["companyFilePath"], json);

            List<Item> newitems=LoadJson();
            populateCombo(newitems);
        }

        private void exportbtn_Click(object sender, EventArgs e)
        {
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
            #region -oldcode2 
            //System.Data.DataTable temDt = Import_Template();
            ////temDt
            //object missing = System.Reflection.Missing.Value;
            //Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            //Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(missing);

            //Microsoft.Office.Interop.Excel.Worksheet newWorksheet;
            //newWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)excel.Worksheets.Add(missing, missing, missing, missing);
            //newWorksheet.Name = "Name of data sheet";

            ////  for first datatable dt1..


            //System.Data.DataTable info = new System.Data.DataTable();
            //info.Columns.AddRange(new DataColumn[4] { 
            //new DataColumn("Hours", typeof(string)),
            //new DataColumn("person name",typeof(string)),
            //new DataColumn("person second name",typeof(string)),
            //new DataColumn("division",typeof(string))
            
            //});

            //for (int i = 0; i < dt.Rows.Count; i++)
            //{
            //    info.Rows.Add(dt.Rows[i]["F6"], dt.Rows[i]["F1"], dt.Rows[i]["F2"], dt.Rows[i]["F3"]);
            //}


           

            //int iRow1 = 4;
            //foreach (DataRow r in info.Rows)
            //{
            //    iRow1++;

            //    for (int i = 1; i < info.Columns.Count + 1; i++)
            //    {
            //        if (iRow1 <=8)
            //        {
            //        }
            //        else
            //        {
            //            excel.Cells[iRow1 + 1, i + 1] = r[i - 1].ToString();
            //        }
            //    }

            //}

            
            //int iCol1 = 1;
            //foreach (DataColumn c in info.Columns)
            //{
                
            //    iCol1++;
            //    excel.Cells[9, iCol1] = c.ColumnName;
            //    excel.Cells.Font.Bold = true;
            //}

            //excel.Cells.Font.Bold = false;


            ////  for  second datatable dt2..
            //int iCol2 = 0;
            //System.Data.DataTable expdt = new System.Data.DataTable();
            //expdt.Columns.AddRange(new DataColumn[2] { 
            //new DataColumn("ColumnName", typeof(string)),
            //new DataColumn("Column Value",typeof(string)) });
            //expdt.Rows.Add("Company Name:", cmpname.Text);
            //expdt.Rows.Add("Description:", description.Text);
            //expdt.Rows.Add("City:", city.Text);
            //expdt.Rows.Add("Region:", region.Text);
            //expdt.Rows.Add("Street:", street.Text);

            ////foreach (DataColumn c in expdt.Columns)
            ////{
            ////    iCol2++;
            ////    excel.Cells[1, iCol2] = c.ColumnName;
            ////}


            //int iRow2 = -1;
            //foreach (DataRow r in expdt.Rows)
            //{
            //    iRow2++;

            //    for (int i = 1; i < expdt.Columns.Count + 1; i++)
            //    {

            //        //if (iRow2 == 1)
            //        //{
            //        //    // Add the header the first time through 
            //        //    //excel.Cells[iRow2, i] = expdt.Columns[i - 1].ColumnName;
            //        //}
                    
            //        excel.Cells[iRow2 + 1, i] = r[i - 1].ToString();
                    
            //    }

            //}
            //if (!Directory.Exists(ConfigurationManager.AppSettings["ExportExcelPath"]))
            //{
            //    Directory.CreateDirectory(ConfigurationManager.AppSettings["ExportExcelPath"]);
            //}
            
            //workbook.SaveAs(ConfigurationManager.AppSettings["ExportExcelPath"]+filen + ".xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, missing, missing, missing, missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, missing, missing, missing, missing, missing);
            //workbook.Close(true, missing, missing);
            //excel.Quit();
#endregion
            
            ExportMethod();
            exportbtn.Enabled = false;
        }

        public Setting LoadJsonSettings()
        { 
            string path = System.Configuration.ConfigurationManager.AppSettings["settingFilePath"];
            if(!File.Exists(path))
            {
                lblError.Text="Setting.json is not found";
                lblError.ForeColor=System.Drawing.Color.Red;
            }

            using (StreamReader r = new StreamReader(path))
            {
                string txt = r.ReadToEnd();
                Setting _setting = JsonConvert.DeserializeObject<Setting>(txt);
                return _setting;
            }
        
        }

        public bool importMethod(Setting _setting)
        {

            bool status = CheckKeyword(_setting);
            if (status == true)
            {
                Microsoft.Office.Interop.Excel._Application excel_app = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook workbook = excel_app.Workbooks.Open(
                 openFileDialog1.FileName,
                 Type.Missing, true, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing);
                int SheetNumber = Int32.Parse(System.Configuration.ConfigurationManager.AppSettings["SheetNumber"]);
                Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[SheetNumber];
                Microsoft.Office.Interop.Excel.Range range = sheet.get_Range(_setting.exportColumnsLocation, Type.Missing);        //"A5:I10"
                SetTitleAndListValues(sheet, 1, 1);
                //range.Columns.ClearFormats();
                //range.Rows.ClearFormats();
                sheet.Columns.ClearFormats();
                sheet.Rows.ClearFormats();

                int Totalcol = range.Columns.Count;
                int Totalrows = range.Rows.Count;
                object[,] range_values = (object[,])range.Value2;
                System.Data.DataTable dtab = new System.Data.DataTable();
                for (int i = 1; i <= Totalrows; i++)
                {

                    for (int j = 1; j <= Totalcol; j++)
                    {
                        object _currentCell = (object)range_values[i, j];


                        if (_currentCell != null)     //&&( _currentCell.ToString() == "name" || _currentCell.ToString() == "surname" || _currentCell.ToString() == "division" || _currentCell.ToString() == "hour")
                        {
                            for (int t = 0; t < _setting.exportColumns.Count; t++)
                            {
                                if (_currentCell.ToString() == _setting.exportColumns[t])
                                {
                                    dtab.Columns.Add(_currentCell.ToString());
                                }
                            }

                        }
                    }

                }

                workbook.Close(false, Type.Missing, Type.Missing);

                if(dtab.Columns.Count < 1)
                {
                    lblError.Text="Columns are not in given locations";
                    lblError.ForeColor=System.Drawing.Color.Red;
                    return false;
                }
                else
                {
                    ImportDataTable = dtab;
                    return true;
                }
                   
            }
            else
            {
                lblError.Text="Keyword is not found in te Given location";
                lblError.ForeColor=System.Drawing.Color.Red;
                return false;
            }

        }

        private void SetTitleAndListValues(Microsoft.Office.Interop.Excel.Worksheet sheet,int row, int col)
        {
            Microsoft.Office.Interop.Excel.Range range;

            // Set the title.
            range = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[row, col];
            

            // Get the values.
            // Find the last cell in the column.
            range = (Microsoft.Office.Interop.Excel.Range)sheet.Columns[col, Type.Missing];
            Microsoft.Office.Interop.Excel.Range last_cell =
                range.get_End(Microsoft.Office.Interop.Excel.XlDirection.xlDown);

            // Get a Range holding the values.
            Microsoft.Office.Interop.Excel.Range first_cell =
                (Microsoft.Office.Interop.Excel.Range)sheet.Cells[row + 1, col];
            Microsoft.Office.Interop.Excel.Range value_range =
                (Microsoft.Office.Interop.Excel.Range)sheet.get_Range(first_cell, last_cell);

            // Get the values.
            object[,] range_values = (object[,])value_range.Value2;

            // Convert this into a 1-dimensional array.
            // Note that the Range's array has lower bounds 1.
            int num_items = range_values.GetUpperBound(0);
            string[] values1 = new string[num_items];
            for (int i = 0; i < num_items; i++)
            {
                values1[i] = (string)range_values[i + 1, 1];
            }

            // Display the values in the ListBox.
            //lst.DataSource = values1;
        }

        public bool CheckKeyword(Setting _setting)
        {
            Microsoft.Office.Interop.Excel._Application excel_app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = excel_app.Workbooks.Open(
             openFileDialog1.FileName,
             Type.Missing, true, Type.Missing, Type.Missing,
             Type.Missing, Type.Missing, Type.Missing, Type.Missing,
             Type.Missing, Type.Missing, Type.Missing, Type.Missing,
             Type.Missing, Type.Missing);
            int SheetNumber=Int32.Parse( System.Configuration.ConfigurationManager.AppSettings["SheetNumber"]);
            Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[SheetNumber];
            Microsoft.Office.Interop.Excel.Range range = sheet.get_Range(_setting.tabPatternLocation, Type.Missing);        //"A5:I10"
            SetTitleAndListValues(sheet, 1, 1);
            sheet.Columns.ClearFormats();
            sheet.Rows.ClearFormats();

            int Totalcol = range.Columns.Count;
            int Totalrows = range.Rows.Count;
            object[,] range_values = (object[,])range.Value2;

            for (int i = 1; i <= Totalrows; i++)
            {

                for (int j = 1; j <= Totalcol; j++)
                {
                    object _currentCell = (object)range_values[i, j];
                    if (_currentCell != null && _currentCell.ToString()==_setting.tabPatternText)
                    {
                        workbook.Close(false, Type.Missing, Type.Missing);
                        return true;
                    }
                }

            }
            workbook.Close(false, Type.Missing, Type.Missing);
            return false;

        }

        public System.Data.DataTable Import_Template()
        {
            #region -oldcode
            //string name = System.Configuration.ConfigurationManager.AppSettings["templateFilePath"];
            //string ext = System.IO.Path.GetExtension(name).ToLower();
            //if (!ext.Equals(".xlsx"))
            //{
            //    lblError.Text = "File format should be xlsx";
            //    lblError.ForeColor = System.Drawing.Color.Red;
            //    return null;
            //}
            //string[] _nameofFile = new string[5];
            //string conStr, sheetName;
            //string header = "NO";
            //conStr = string.Empty;
            //switch (ext)
            //{

            //    case ".xls": //Excel 97-03
            //        conStr = string.Format(Excel03ConString, name, header);
            //        break;

            //    case ".xlsx": //Excel 07
            //        conStr = string.Format(Excel07ConString, name, header);
            //        break;
            //}

            //using (OleDbConnection con = new OleDbConnection(conStr))
            //{
            //    using (OleDbCommand cmd = new OleDbCommand())
            //    {
            //        cmd.Connection = con;
            //        con.Open();
            //        System.Data.DataTable dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            //        sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
            //        con.Close();
            //    }
            //}
            //System.Data.DataTable templateDTable = new System.Data.DataTable();
            //using (OleDbConnection con = new OleDbConnection(conStr))
            //{
            //    using (OleDbCommand cmd = new OleDbCommand())
            //    {
            //        using (OleDbDataAdapter oda = new OleDbDataAdapter())
            //        {
            //            cmd.CommandText = "SELECT * From [" + sheetName + "]";
            //            cmd.Connection = con;
            //            con.Open();
            //            oda.SelectCommand = cmd;
            //            oda.Fill(templateDTable);
            //            con.Close();
            //            //exportbtn.Enabled = true;
            //            return templateDTable;
            //        }
            //    }
            //}
            #endregion
            Microsoft.Office.Interop.Excel._Application excel_app = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook workbook = excel_app.Workbooks.Open(
                 openFileDialog1.FileName,
                 Type.Missing, true, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing);
                Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[4];
                int Totalcol = sheet.Columns.Count;
                int Totalrows = sheet.Rows.Count;
                Microsoft.Office.Interop.Excel.Range range = sheet.get_Range(Totalcol,Totalrows);        //"A5:I10"
                SetTitleAndListValues(sheet, 1, 1);
                //range.Columns.ClearFormats();
                //range.Rows.ClearFormats();
                

                int Tcol = range.Columns.Count;
                int Trows = range.Rows.Count;
                
                object[,] range_values = (object[,])range.Value2;
                System.Data.DataTable dtab = new System.Data.DataTable();
                for (int i = 1; i <= Totalrows; i++)
                {

                    for (int j = 1; j <= Totalcol; j++)
                    {
                        object _currentCell = (object)range_values[i, j];


                        if (_currentCell != null)     //&&( _currentCell.ToString() == "name" || _currentCell.ToString() == "surname" || _currentCell.ToString() == "division" || _currentCell.ToString() == "hour")
                        {
                            for (int t = 0; t < ImportDataTable.Columns.Count; t++)
                            {
                                if (_currentCell.ToString() == ImportDataTable.Columns[t].ColumnName)
                                {
                                    dtab.Columns.Add(_currentCell.ToString());
                                }
                            }

                        }
                    }

                }

                ImportDataTable = dtab;

                workbook.Close(false, Type.Missing, Type.Missing);
            
            return ImportDataTable;// wrong code
        }
        
        public void ExportMethod()
        {
            System.Data.DataTable dtExp=ImportDataTable;
            string templateFile=System.Configuration.ConfigurationManager.AppSettings["templateFilePath"];//@"C:\Users\muhammad.habib\Downloads\new folder\template1.xlsx"
             if (!File.Exists(templateFile))
            {
                lblError.Text = "template.xlsx file does not exist";
                lblError.ForeColor = System.Drawing.Color.Red;
                return;   
            }
            Microsoft.Office.Interop.Excel._Application excel_app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = excel_app.Workbooks.Open(
            templateFile,
             Type.Missing, true, Type.Missing, Type.Missing,
             Type.Missing, Type.Missing, Type.Missing, Type.Missing,
             Type.Missing, Type.Missing, Type.Missing, Type.Missing,
             Type.Missing, Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];

            #region -making new sheet
            Microsoft.Office.Interop.Excel.Worksheet newWorksheet;
            newWorksheet = sheet;
            #endregion
            
            Microsoft.Office.Interop.Excel.Range last = sheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Microsoft.Office.Interop.Excel.Range range = newWorksheet.get_Range("A1",last);
            SetTitleAndListValues(newWorksheet, 1, 1);
            int Sheetcol = range.Columns.Count;
            int Sheetrow = range.Rows.Count;
            
            object[,] range_values = (object[,])range.Value2;
            for (int r = 1; r <= Sheetrow; r++)
            {
                for (int c = 1; c <= Sheetcol; c++)
                {
                    object CurrentVal= (object)range_values[r,c] ;
                    if ( CurrentVal!=null)
                    {
                        if (CurrentVal.ToString() == "{companyName}")
                        {
                            newWorksheet.Cells[r, c] = cmpname.Text;
                        }
                        else if (CurrentVal.ToString() == "{description}")
                        {
                            newWorksheet.Cells[r, c] = description.Text;
                        }
                        else if (CurrentVal.ToString() == "{city}")
                        {
                            newWorksheet.Cells[r, c] = city.Text;
                        }
                        else if (CurrentVal.ToString() == "{street}")
                        {
                            newWorksheet.Cells[r, c] = street.Text;
                        }
                        else if (CurrentVal.ToString() == "{region}")
                        {
                            newWorksheet.Cells[r, c] = region.Text;
                        }

                        for (int colsdt = 0; colsdt < dtExp.Columns.Count; colsdt++)
                        {
                            string columnNameofExpdt="{" + dtExp.Columns[colsdt].ColumnName + "}";
                            if (CurrentVal.ToString() ==columnNameofExpdt)
                            { 
                                //newWorksheet.Cells[r, c]=dtExp.Rows[]
                                for (int rr = 1; rr < dtExp.Rows.Count; rr++)
                                {
                                    if (rr == 1)
                                    {
                                        newWorksheet.Cells[r, c] = dtExp.Rows[rr][colsdt].ToString();
                                    }
                                    else
                                    {
                                        newWorksheet.Cells[r+rr-1, c] = dtExp.Rows[rr][colsdt].ToString();
                                    }
                                }
                            }
                        }
                      
                    }
                }
            }
            try
            {
                string expPath = System.Configuration.ConfigurationManager.AppSettings["ExportExcelPath"];
                if (!Directory.Exists(expPath))
                {
                    Directory.CreateDirectory(expPath);
                }
                workbook.SaveAs(expPath+"/" + filen + ".xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                lblError.Text = "File Exported Successfully";
                lblError.ForeColor = System.Drawing.Color.Green;
            }
            catch
            {
                lblError.Text = "Export Failed";
                lblError.ForeColor = System.Drawing.Color.Red;
            }
               
            workbook.Close(true, Type.Missing, Type.Missing);
            //excel.Quit(); 
        }
        
        
    
    
    
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

    public class Setting
    {
        public string tabPatternText;
        public string tabPatternLocation;
        public string exportColumnsLocation;
        public List<string> exportColumns;
    }

    public static class DataTableExtensions
    {
        public static System.Data.DataTable SetColumnsOrder(this System.Data.DataTable table, params String[] columnNames)
        {
            int columnIndex = 0;
            foreach (var columnName in columnNames)
            {
                table.Columns[columnName].SetOrdinal(columnIndex);
                columnIndex++;
            }
            return table;
        }

    }






