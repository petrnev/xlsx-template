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
            comboBox1.SelectedIndex = 1;
           
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
                lblError.Text = "File Uploading..." + "\n" + lblError.Text;
                //lblError.ForeColor = System.Drawing.Color.Green;  
            }
            else
            {
                return;
            }

           
            string name = openFileDialog1.FileName;
            string ext = System.IO.Path.GetExtension(openFileDialog1.FileName).ToLower();   
            if (!ext.Equals(".xlsx"))
            {
                lblError.Text = lblError.Text + "\nFile format should be xlsx";
                //lblError.ForeColor = System.Drawing.Color.Red;
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
            int Totalsheets;
            using (OleDbConnection con = new OleDbConnection(conStr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    cmd.Connection = con;
                    con.Open();
                    System.Data.DataTable dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    //int SheetNumber = Int32.Parse(System.Configuration.ConfigurationManager.AppSettings["SheetNumber"]);
                    Totalsheets=dtExcelSchema.Rows.Count;
                    //sheetName = dtExcelSchema.Rows[SheetNumber-1]["TABLE_NAME"].ToString();
                    con.Close();
                }
            }
           // DataTable dt = new DataTable();
            dt = new System.Data.DataTable();
            System.Data.DataSet ds= new DataSet();
            int size=0;
            for(int SheetNo=0;SheetNo<Totalsheets;SheetNo++)
            {
                System.Data.DataTable SheetDt= new System.Data.DataTable();
                using (OleDbConnection con = new OleDbConnection(conStr))
                {
                    using (OleDbCommand cmd = new OleDbCommand())
                    {
                        con.Open();
                        System.Data.DataTable dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        sheetName = dtExcelSchema.Rows[SheetNo]["TABLE_NAME"].ToString();
                        con.Close();
                        using (OleDbDataAdapter oda = new OleDbDataAdapter())
                        {

                            string[] limit= lst.exportColumnsLocation.Split(':');
                            cmd.CommandText = "SELECT * From [" + sheetName + "]";
                            cmd.Connection = con;
                            con.Open();
                            oda.SelectCommand = cmd;
                            oda.Fill(SheetDt);
                            con.Close();
                            //exportbtn.Enabled = true;
                            //Populate DataGridView.
                            //dataGridView1.DataSource = dt;
                        }
                    }
                }
                if(SheetDt!=null)
                {
                    ds.Tables.Add(SheetDt);
                    size=size+SheetDt.Rows.Count;
                }
            }

            if (!importMethod(lst))//if (!importMethod(lst,Totalsheets))
            {
                return;
            }

            ds = RemoveTables(ds, lst);
            #region -new code for import
            string[] valname=new string[size];
            int Position = 0;
            int[] posIndex = new int[ds.Tables.Count];
            int addedrecords = 0;
            for (int TotalTables = 0; TotalTables < ds.Tables.Count; TotalTables++)
            {
               
                if(TotalTables==0)
                {
                   posIndex[TotalTables] = Position; 
                   Position = populateDataTables(ds.Tables[TotalTables], valname, Position);
                   addedrecords = Position - 1;
                   //lblError.Text = lblError.Text + "\n" + addedrecords + " Records are imported from Sheet" + ds.Tables[TotalTables].TableName.Substring(ds.Tables[TotalTables].TableName.Length - 1);
                   //lblError.ForeColor = System.Drawing.Color.Green;
                   lblError.Text = addedrecords + " Records are imported from Sheet" + ds.Tables[TotalTables].TableName.Substring(ds.Tables[TotalTables].TableName.Length - 1) + "\n" + lblError.Text;
                }
                else
                {
                    posIndex[TotalTables] = Position;
                    Position = populateDataTables(ds.Tables[TotalTables], valname, Position);
                    addedrecords = Position - addedrecords-TotalTables-1;
                    //lblError.Text = lblError.Text + "\n" + addedrecords + " Records are imported from Sheet" + ds.Tables[TotalTables].TableName.Substring(ds.Tables[TotalTables].TableName.Length - 1);
                    //lblError.ForeColor = System.Drawing.Color.Green;
                    lblError.Text = addedrecords + " Records are imported from Sheet" + ds.Tables[TotalTables].TableName.Substring(ds.Tables[TotalTables].TableName.Length - 1) + "\n" + lblError.Text;
                    addedrecords = Position - TotalTables - 1;
                }
            }

            System.Data.DataTable _impdt = ImportDataTable;
            #region -oldcode
            //string[] valname = new string[dt.Rows.Count];
            //for (int a = 0; a < dt.Rows.Count; a++)
            //{
            //    for (int b = 0; b < dt.Columns.Count; b++)
            //    {

            //        for (int f = 0; f < _impdt.Columns.Count; f++)
            //        {

            //            if (dt.Rows[a][b].ToString() == _impdt.Columns[f].ColumnName)
            //            {

            //                //int remainig = a;
            //                //string[] valname=new  string[dt.Rows.Count-a];
            //                int d = 0;
            //                for (int c = a; c < dt.Rows.Count; c++)
            //                {
            //                    if (valname[d] == null)
            //                    {
            //                        valname[d] = dt.Rows[c][b].ToString();
            //                    }
            //                    else
            //                    {
            //                        valname[d] = valname[d] + "," + dt.Rows[c][b].ToString();
            //                    }
            //                    d++;
            //                }

            //            }

            //        }

            //        #region -oldcode
            //        //if (dt.Rows[a][b].ToString() == "name" || dt.Rows[a][b].ToString() == "surname" || dt.Rows[a][b].ToString() == "Divisions")
            //        //{

            //        //    int remainig = a;
            //        //    //string[] valname=new  string[dt.Rows.Count-a];
            //        //    int d = 0;
            //        //    for (int c = a + 1; c < dt.Rows.Count; c++)
            //        //    {
            //        //        if (valname[d] == null)
            //        //        {
            //        //            valname[d] = dt.Rows[c][b].ToString();
            //        //        }
            //        //        else
            //        //        {
            //        //            valname[d] = valname[d] + "," + dt.Rows[c][b].ToString();
            //        //        }
            //        //        d++;
            //        //    }

            //        //}
            //        #endregion

            //    }
            //}
            #endregion
            Setting _setting = LoadJsonSettings();
            string[] newarray = valname.Where(c => c != null).ToArray();

            for (int h = 0; h < newarray.Count(); h++)
            {
                string[] val = newarray[h].Split('#');
                //if (h == 0)
                //{
                //    _impdt = DataTableExtensions.SetColumnsOrder(_impdt, val);
                //    //DataTableExtensions dte = new DataTableExtensions();
                //    //_impdt = dte.SetColumnsOrder(_impdt, val);
                //}
                for (int match = 0; match < posIndex.Count(); match++)
                {
                    if (h == posIndex[match])
                    {
                        //lblError.Text=lblError.Text+""
                        _impdt = DataTableExtensions.SetColumnsOrder(_impdt, val);
                    }
                
                }

                DataRow row = _impdt.NewRow();

                for (int g = 0; g < _impdt.Columns.Count; g++)
                {
                    row[g] = val[g];
                }

                _impdt.Rows.Add(row);
            }


            for (int rem = 0; rem < posIndex.Count(); rem++)
            {
                _impdt.Rows.RemoveAt(posIndex[rem]-rem);
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
            exportbtn.Enabled = true;
            lblError.Text = "File Uploaded"+"\n"+lblError.Text;
            //lblError.ForeColor = System.Drawing.Color.Green;
          
        }

        public DataSet RemoveTables(DataSet ds,Setting _setting)
        {
            Microsoft.Office.Interop.Excel._Application excel_app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = excel_app.Workbooks.Open(
             openFileDialog1.FileName,
             Type.Missing, true, Type.Missing, Type.Missing,
             Type.Missing, Type.Missing, Type.Missing, Type.Missing,
             Type.Missing, Type.Missing, Type.Missing, Type.Missing,
             Type.Missing, Type.Missing);
            DataSet _newds=ds;
            int TablesCounters=ds.Tables.Count;
            int TablesCount = workbook.Sheets.Count;
            int iCounter = 0;
            for (int TableNumber = 0; TableNumber < TablesCount; TableNumber++)
            {
                if (!CheckKeyword(_setting, TableNumber,excel_app,workbook))
                {
                    _newds.Tables.RemoveAt(TableNumber-iCounter);
                    iCounter++;
                }
            }

            #region -- Checking export columns location
            int Jcounter = 0;
            for (int TableNumber = 0; TableNumber < ds.Tables.Count; TableNumber++)
            {
                int SheetNumber = Int32.Parse(ds.Tables[TableNumber].TableName.Substring(ds.Tables[TableNumber].TableName.Length-1));
                Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[SheetNumber];
                Microsoft.Office.Interop.Excel.Range range = sheet.get_Range(_setting.exportColumnsLocation, Type.Missing);        //"A5:I10"
                SetTitleAndListValues(sheet, 1, 1);
                //range.Columns.ClearFormats();
                //range.Rows.ClearFormats();
                sheet.Columns.ClearFormats();
                sheet.Rows.ClearFormats();
                System.Data.DataTable dtab = new System.Data.DataTable();
                int Totalcol = range.Columns.Count;
                int Totalrows = range.Rows.Count;
                object[,] range_values = (object[,])range.Value2;

                for (int i = 1; i <= Totalrows; i++)
                {

                    for (int j = 1; j <= Totalcol; j++)
                    {
                        object _currentCell = (object)range_values[i, j];


                        if (_currentCell != null)     //&&( _currentCell.ToString() == "name" || _currentCell.ToString() == "surname" || _currentCell.ToString() == "division" || _currentCell.ToString() == "hour")
                        {
                            for (int t = 0; t < _setting.exportColumns.Count; t++)
                            {
                                if (_currentCell.ToString().ToLower() == _setting.exportColumns[t].ToLower())
                                {
                                    dtab.Columns.Add(_currentCell.ToString());
                                }
                            }

                        }
                    }

                }

            #region oldcode
		                //workbook.Close(false, Type.Missing, Type.Missing);

                //if (dtab.Columns.Count < 1)
                //{
                //    lblError.Text = "Columns are not in given locations";
                //    lblError.ForeColor = System.Drawing.Color.Red;
                //    return false;
                //}
                //else
                //{
                //    ImportDataTable = dtab;
                //     

                //}
            #endregion
                if (dtab.Columns.Count <1)
                {
                    ds.Tables.RemoveAt(TableNumber-Jcounter);
                    Jcounter++;
                }
            }
            #endregion
            workbook.Close(false, Type.Missing, Type.Missing);
            return ds;
        }

        public List<Item> LoadJson()
        {
            string path = System.Configuration.ConfigurationManager.AppSettings["companyFilePath"];
            if (!File.Exists(path))
            {
                lblError.Text = "companies.Json file does not exist" + "\n" + lblError.Text;
                //lblError.ForeColor = System.Drawing.Color.Red;
                return null;   
            }

            //lblError.Text = string.Empty;
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
                lblError.Text = lblError.Text+"\ncompanies.Json file does not exist"+"\n"+lblError.Text;
                //lblError.ForeColor = System.Drawing.Color.Red;
                //return null;
            }
           
            //lblError.Text = string.Empty;
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
                lblError.Text = "You cannot save data now"+"\n"+lblError.Text;
                //lblError.ForeColor = System.Drawing.Color.Red;
                return;
            }

            bool _isupdate = false;
            int c=items.Count;
            c=c+1;
                for (int i = 0; i < items.Count; i++)
                {
                    if (items[i].companyname == cmpname.Text)
                    {
                        DialogResult result1 = MessageBox.Show("Do you want to update current company information?","Update Alert",MessageBoxButtons.YesNo);
                        if (result1 == DialogResult.Yes)
                        {
                            c = Int32.Parse(items[i].id);
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
                        _isupdate = true;
                        break;
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
                lblError.Text = "You cannot delete data now"+"\n"+lblError.Text;
                //lblError.ForeColor = System.Drawing.Color.Red;
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
                lblError.Text="Setting.json is not found"+"\n"+lblError.Text;
                //lblError.ForeColor=System.Drawing.Color.Red;
                return null;
            }

            using (StreamReader r = new StreamReader(path))
            {
                string txt = r.ReadToEnd();
                Setting _setting = JsonConvert.DeserializeObject<Setting>(txt);
                return _setting;
            }
        
        }

        //public bool importMethod(Setting _setting,int TotalSheet)
        //{

        //    bool status = CheckKeyword(_setting);
        //    if (status == true)
        //    {
        //        Microsoft.Office.Interop.Excel._Application excel_app = new Microsoft.Office.Interop.Excel.Application();
        //        Microsoft.Office.Interop.Excel.Workbook workbook = excel_app.Workbooks.Open(
        //         openFileDialog1.FileName,
        //         Type.Missing, true, Type.Missing, Type.Missing,
        //         Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //         Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //         Type.Missing, Type.Missing);
        //        int SheetNumber = Int32.Parse(System.Configuration.ConfigurationManager.AppSettings["SheetNumber"]);
        //        Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[SheetNumber];
        //        Microsoft.Office.Interop.Excel.Range range = sheet.get_Range(_setting.exportColumnsLocation, Type.Missing);        //"A5:I10"
        //        SetTitleAndListValues(sheet, 1, 1);
        //        //range.Columns.ClearFormats();
        //        //range.Rows.ClearFormats();
        //        sheet.Columns.ClearFormats();
        //        sheet.Rows.ClearFormats();

        //        int Totalcol = range.Columns.Count;
        //        int Totalrows = range.Rows.Count;
        //        object[,] range_values = (object[,])range.Value2;
        //        System.Data.DataTable dtab = new System.Data.DataTable();
        //        for (int i = 1; i <= Totalrows; i++)
        //        {

        //            for (int j = 1; j <= Totalcol; j++)
        //            {
        //                object _currentCell = (object)range_values[i, j];


        //                if (_currentCell != null)     //&&( _currentCell.ToString() == "name" || _currentCell.ToString() == "surname" || _currentCell.ToString() == "division" || _currentCell.ToString() == "hour")
        //                {
        //                    for (int t = 0; t < _setting.exportColumns.Count; t++)
        //                    {
        //                        if (_currentCell.ToString().ToLower() == _setting.exportColumns[t].ToLower())
        //                        {
        //                            dtab.Columns.Add(_currentCell.ToString());
        //                        }
        //                    }

        //                }
        //            }

        //        }

        //        workbook.Close(false, Type.Missing, Type.Missing);

        //        if(dtab.Columns.Count < 1)
        //        {
        //            lblError.Text="Columns are not in given locations";
        //            lblError.ForeColor=System.Drawing.Color.Red;
        //            return false;
        //        }
        //        else
        //        {
        //            ImportDataTable = dtab;
        //            return true;
        //        }
                   
        //    }
        //    else
        //    {
        //        lblError.Text="Keyword is not found in the Given location";
        //        lblError.ForeColor=System.Drawing.Color.Red;
        //        return false;
        //    }

        //}

        public bool importMethod(Setting _setting)
        {
            System.Data.DataTable dtab = new System.Data.DataTable();
            bool checkedAllSheets = false;
            Microsoft.Office.Interop.Excel._Application excel_app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = excel_app.Workbooks.Open(
             openFileDialog1.FileName,
             Type.Missing, true, Type.Missing, Type.Missing,
             Type.Missing, Type.Missing, Type.Missing, Type.Missing,
             Type.Missing, Type.Missing, Type.Missing, Type.Missing,
             Type.Missing, Type.Missing);
            int TotalSheet = workbook.Sheets.Count;
            for (int SheetCounter = 0; SheetCounter < TotalSheet; SheetCounter++)
            {
                bool status = CheckKeyword(_setting,SheetCounter,excel_app,workbook);
                if (status == true)
                {
                    
                    //int SheetNumber = Int32.Parse(System.Configuration.ConfigurationManager.AppSettings["SheetNumber"]);
                    Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[SheetCounter+1];
                    Microsoft.Office.Interop.Excel.Range range = sheet.get_Range(_setting.exportColumnsLocation, Type.Missing);        //"A5:I10"
                    SetTitleAndListValues(sheet, 1, 1);
                    //range.Columns.ClearFormats();
                    //range.Rows.ClearFormats();
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


                            if (_currentCell != null)     //&&( _currentCell.ToString() == "name" || _currentCell.ToString() == "surname" || _currentCell.ToString() == "division" || _currentCell.ToString() == "hour")
                            {
                                for (int t = 0; t < _setting.exportColumns.Count; t++)
                                {
                                    if (_currentCell.ToString().ToLower() == _setting.exportColumns[t].ToLower())
                                    {
                                        dtab.Columns.Add(_currentCell.ToString());
                                    }
                                }

                            }
                        }

                    }

                    //workbook.Close(false, Type.Missing, Type.Missing);

                    //if (dtab.Columns.Count < 1)
                    //{
                    //    lblError.Text = "Columns are not in given locations";
                    //    lblError.ForeColor = System.Drawing.Color.Red;
                    //    return false;
                    //}
                    //else
                    //{
                    //    ImportDataTable = dtab;
                    //    return true;
                    //}
                    if (dtab.Columns.Count > 0)
                    {
                        checkedAllSheets = true;
                        break;
                    }
                }
                else
                {
                    //lblError.Text = "Keyword is not found in the Given location";
                    //lblError.ForeColor = System.Drawing.Color.Red;
                    //return false;
                    checkedAllSheets = false;
                }
            }

            workbook.Close(false, Type.Missing, Type.Missing);

            if (checkedAllSheets)
            {
                if (dtab.Columns.Count < 1)
                {
                    lblError.Text = "Columns are not in given locations"+"\n"+lblError.Text;
                    //lblError.ForeColor = System.Drawing.Color.Red;
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
                lblError.Text = "Keyword is not found in the Given location"+"\n"+lblError.Text;
                //lblError.ForeColor = System.Drawing.Color.Red;
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
            //string[] values1 = new string[num_items];
            object[] values1 = new object[num_items];
            for (int i = 0; i < num_items; i++)
            {
                values1[i] = (object)range_values[i + 1, 1];
            }

            // Display the values in the ListBox.
            //lst.DataSource = values1;
        }

        //public bool CheckKeyword(Setting _setting)
        //{
        //    Microsoft.Office.Interop.Excel._Application excel_app = new Microsoft.Office.Interop.Excel.Application();
        //    Microsoft.Office.Interop.Excel.Workbook workbook = excel_app.Workbooks.Open(
        //     openFileDialog1.FileName,
        //     Type.Missing, true, Type.Missing, Type.Missing,
        //     Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //     Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //     Type.Missing, Type.Missing);
        //    int SheetNumber=Int32.Parse( System.Configuration.ConfigurationManager.AppSettings["SheetNumber"]);
        //    Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[SheetNumber];
        //    Microsoft.Office.Interop.Excel.Range range = sheet.get_Range(_setting.tabPatternLocation, Type.Missing);        //"A5:I10"
        //    SetTitleAndListValues(sheet, 1, 1);
        //    sheet.Columns.ClearFormats();
        //    sheet.Rows.ClearFormats();

        //    int Totalcol = range.Columns.Count;
        //    int Totalrows = range.Rows.Count;
        //    object[,] range_values = (object[,])range.Value2;

        //    for (int i = 1; i <= Totalrows; i++)
        //    {

        //        for (int j = 1; j <= Totalcol; j++)
        //        {
        //            object _currentCell = (object)range_values[i, j];
        //            if (_currentCell != null && _currentCell.ToString()==_setting.tabPatternText)
        //            {
        //                workbook.Close(false, Type.Missing, Type.Missing);
        //                return true;
        //            }
        //        }

        //    }
        //    workbook.Close(false, Type.Missing, Type.Missing);
        //    return false;

        //}

        public bool CheckKeyword(Setting _setting, int SheetNumber, Microsoft.Office.Interop.Excel._Application excel_app, Microsoft.Office.Interop.Excel.Workbook workbook)
        {
            //Microsoft.Office.Interop.Excel._Application excel_app = new Microsoft.Office.Interop.Excel.Application();
            //Microsoft.Office.Interop.Excel.Workbook workbook = excel_app.Workbooks.Open(
            // openFileDialog1.FileName,
            // Type.Missing, true, Type.Missing, Type.Missing,
            // Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            // Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            // Type.Missing, Type.Missing);


            //int SheetNumber = Int32.Parse(System.Configuration.ConfigurationManager.AppSettings["SheetNumber"]);
            Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[SheetNumber+1];
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
                    if (_currentCell != null && _currentCell.ToString() == _setting.tabPatternText)
                    {
                        //workbook.Close(false, Type.Missing, Type.Missing);
                        return true;
                    }
                }

            }
            //workbook.Close(false, Type.Missing, Type.Missing);
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

        public System.Data.DataTable DeleteEmptyRows(System.Data.DataTable EmptDt)
        {
            
            for (int i = 0; i < EmptDt.Rows.Count; i++)
            {
                String valuesarr = String.Empty;
                List<object> lst = EmptDt.Rows[i].ItemArray.ToList();
                foreach (Object s in lst)
                {
                    valuesarr += s.ToString();
                }

                if (String.IsNullOrEmpty(valuesarr))
                    EmptDt.Rows.RemoveAt(i);
            }
            return EmptDt;
        }

        public void ExportMethod()
        {
            System.Data.DataTable dtExp=ImportDataTable;
            dtExp = DeleteEmptyRows(dtExp);
            string templateFile=System.Configuration.ConfigurationManager.AppSettings["templateFilePath"];//@"C:\Users\muhammad.habib\Downloads\new folder\template1.xlsx"
             if (!File.Exists(templateFile))
            {
                lblError.Text = "\ntemplate.xlsx file does not exist" + "\n" + lblError.Text;
                //lblError.ForeColor = System.Drawing.Color.Red;
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
                        if (CurrentVal.ToString().ToLower() == "{companyname}")
                        {
                            newWorksheet.Cells[r, c] = cmpname.Text;
                        }
                        else if (CurrentVal.ToString().ToLower() == "{description}")
                        {
                            newWorksheet.Cells[r, c] = description.Text;
                        }
                        else if (CurrentVal.ToString().ToLower() == "{city}")
                        {
                            newWorksheet.Cells[r, c] = city.Text;
                        }
                        else if (CurrentVal.ToString().ToLower() == "{street}")
                        {
                            newWorksheet.Cells[r, c] = street.Text;
                        }
                        else if (CurrentVal.ToString().ToLower() == "{region}")
                        {
                            newWorksheet.Cells[r, c] = region.Text;
                        }

                        for (int colsdt = 0; colsdt < dtExp.Columns.Count; colsdt++)
                        {
                            string columnNameofExpdt="{" + dtExp.Columns[colsdt].ColumnName + "}";
                            if (CurrentVal.ToString().ToLower() ==columnNameofExpdt.ToLower())
                            { 
                                //newWorksheet.Cells[r, c]=dtExp.Rows[]
                                for (int rr = 0; rr < dtExp.Rows.Count; rr++)
                                {
                                    //if (rr == 1)
                                    //{
                                    //    newWorksheet.Cells[r, c] = dtExp.Rows[rr][colsdt].ToString();
                                    //}
                                    //else
                                    //{
                                        newWorksheet.Cells[r+rr, c] = dtExp.Rows[rr][colsdt].ToString();
                                    //}
                                }
                            }
                        }
                      
                    }
                }
            }
            //try
            //{
                string expPath = System.Configuration.ConfigurationManager.AppSettings["ExportExcelPath"];
                if (!Directory.Exists(expPath))
                {
                    Directory.CreateDirectory(expPath);
                }
                try
                {
                    workbook.SaveAs(expPath + "/" + filen + ".xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }
                catch
                {
                   
                }
               
            workbook.Close(true, Type.Missing, Type.Missing);
            //excel.Quit(); 
            #region MyRegion
            //string FileCreated=expPath+"/" + filen + ".xlsx";
            //if (File.Exists(FileCreated))
            //{
            //    lblError.Text = filen + ".xlsx Exported Successfully";
            //    lblError.ForeColor = System.Drawing.Color.Red;
            //}
            //else
            //{
            //    lblError.Text = filen + ".xlsx Exported Successfully";
            //    lblError.ForeColor = System.Drawing.Color.Green;
            //} 
            #endregion
            lblError.Text = "File Exported Successfully" + "\n" + lblError.Text;
            //lblError.ForeColor = System.Drawing.Color.Green;
        }

        public int populateDataTables(System.Data.DataTable dt,string[] valname,int pos)
        {
            System.Data.DataTable _impdt = ImportDataTable;
            //string[] valname = new string[dt.Rows.Count];
            for (int a = 0; a < dt.Rows.Count; a++)
            {
                for (int b = 0; b < dt.Columns.Count; b++)
                {

                    for (int f = 0; f < _impdt.Columns.Count; f++)
                    {

                        if (dt.Rows[a][b].ToString().ToLower() == _impdt.Columns[f].ColumnName.ToLower())
                        {

                            //int remainig = a;
                            //string[] valname=new  string[dt.Rows.Count-a];
                            //int d = 0;
                            int d = pos;
                            for (int c = a; c < dt.Rows.Count; c++)
                            {
                                if (valname[d] == null)
                                {
                                    valname[d] = dt.Rows[c][b].ToString();
                                }
                                else
                                {
                                    valname[d] = valname[d] + "#" + dt.Rows[c][b].ToString();
                                }
                                d++;
                            }

                        }

                    }

                }
            }

            var CurrentIndexVar = valname.Select((day, index) => new { Day = day, Index = index }).Where(x => x.Day==null).FirstOrDefault();
            int CurrentIndex = CurrentIndexVar.Index;
            return CurrentIndex;
        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

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






