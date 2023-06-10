using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;

namespace ToJson_Winform
{
    public partial class Form1 : Form
    {
        private string loadFolderName;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void folderAndFileBrowsing(object sender, EventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = Application.StartupPath; // 현재 자신의 위치를 디폴트로 설정함.
            dialog.IsFolderPicker = true; // 폴더 선택도 가능케함.
            
            if(dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                DirectoryInfo directoryInfo = new DirectoryInfo(dialog.FileName);

                FileAttributes checkAtt = File.GetAttributes(directoryInfo.FullName);

                if(checkAtt == FileAttributes.Directory) // 선택한 경로가 폴더디렉토리 일경우.
                {

                }
                else
                {

                }
            }
            else
            {
                Console.WriteLine("exception");
            }
        }
    }
}
