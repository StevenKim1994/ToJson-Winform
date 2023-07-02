using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data;
using System.IO;
using System.Windows.Forms;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace ToJson_Winform
{
    public partial class Form1 : Form
    {
        private string loadFolderName;
        private List<FileInfo> xlslFileList;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            xlslFileList = new List<FileInfo>();
        }

        private static string GetConnectionString(string file)
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            string extension = file.Split('.').Last();

            if (extension == "xls")
            {
                //Excel 2003 and Older
                props["Provider"] = "Microsoft.Jet.OLEDB.4.0";
                props["Extended Properties"] = "Excel 8.0";
            }
            else if (extension == "xlsx")
            {
                //Excel 2007, 2010, 2012, 2013
                props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";
                props["Extended Properties"] = "Excel 12.0 XML";
            }
            else
                throw new Exception(string.Format("error file: {0}", file));

            props["Data Source"] = file;

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }
        private static DataSet GetDataSetFromExcelFile(string file)
        {
            DataSet ds = new DataSet();

            string connectionString = GetConnectionString(file);

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                // Get all Sheets in Excel File
                DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                // Loop through all Sheets to get data
                foreach (DataRow dr in dtSheet.Rows)
                {
                    string sheetName = dr["TABLE_NAME"].ToString();

                    if (!sheetName.EndsWith("$"))
                        continue;

                    // Get all rows from the Sheet
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";

                    DataTable dt = new DataTable();
                    dt.TableName = sheetName;

                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);

                    ds.Tables.Add(dt);
                }

                cmd = null;
                conn.Close();
            }

            return ds;
        }

        private void tableFileBrowsing(object _sender, System.EventArgs _exception)
        {
            string fileKeyword = "Excel (*.xlsl) | *.xlsl";
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.DefaultExt = fileKeyword;
           // dialog.Filter = fileKeyword;
            try
            {
                if(dialog.ShowDialog() == DialogResult.OK)
                {
                    FileInfo fileInfo = new FileInfo(dialog.FileName);
                    Console.WriteLine("LoadFile Name : " + fileInfo.FullName);

                    DataSet dataSet = GetDataSetFromExcelFile(fileInfo.FullName);
                    Console.WriteLine(string.Format("reading file: {0}", fileInfo));
                    Console.WriteLine(string.Format("coloums: {0}", dataSet.Tables[0].Columns.Count));
                    Console.WriteLine(string.Format("rows: {0}", dataSet.Tables[0].Rows.Count));
                }
            }
            catch(Exception _e)
            {
                DialogResult msgBoxResult = MessageBox.Show(_e.ToString(), "잘못된 경로 지정", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (msgBoxResult == DialogResult.OK)
                {
                    // 다시 시도
                    tableFileBrowsing(_sender, _exception);
                }
            }

        }

        private void folderAndFileBrowsing(object _sender, EventArgs _exception)
        {
            string filterKeyword = "Excel (*.xlsl) | *.xlsl";
            CommonOpenFileDialog dialog = new CommonOpenFileDialog("FolderSelect");
            dialog.AllowNonFileSystemItems = true;
            dialog.Multiselect = true;
            dialog.IsFolderPicker = true; // 폴더 선택도 가능케함.
            dialog.DefaultFileName = "변환 시킬 테이블이 존재하는 폴더을 입력하세요.";
            try
            {
                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    xlslFileList.Clear();
                    DirectoryInfo directoryInfo = new DirectoryInfo(dialog.FileName);

                    Console.WriteLine("LoadFolder Name : " + directoryInfo.FullName);

                    // TODO :: 해당 경로에서 xlsl파일이 아닌 경우 불러오지 않도록 예외처리 해야함.

                    foreach (FileInfo file in directoryInfo.GetFiles())
                    {
                        Console.WriteLine(string.Format("Folder In Element: {0}", file.Name));

                        DataSet dataSet = GetDataSetFromExcelFile(file.FullName);
                        Console.WriteLine(string.Format("reading file: {0}", file));
                        Console.WriteLine(string.Format("coloums: {0}", dataSet.Tables[0].Columns.Count));
                        Console.WriteLine(string.Format("rows: {0}", dataSet.Tables[0].Rows.Count));
                    }
                }
                else
                {
                    Console.WriteLine("Exception");
                }
            }
            catch (Exception _e)
            {
                DialogResult msgBoxResult = MessageBox.Show(_e.ToString(), "잘못된 경로 지정", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if(msgBoxResult == DialogResult.OK)
                {
                    // 다시 시도
                    folderAndFileBrowsing(_sender, _exception);
                }
            }
        }
    }
}
