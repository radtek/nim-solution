using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NIM.CertificationGenerator;
using NIM.CertificationGenerator.RadiationThermomater;
using System.Reflection;
using System.Diagnostics;
using System.IO;


namespace NIM_Addin
{
    public partial class UserControl1 : UserControl
    {



        public UserControl1()
        {
            InitializeComponent();
            this.lblVersion.Text = this.GetCurrentApplicationVersion();
            IFilePathManagerProvider.PathProvider = new FilePathManager();
        }


        private string GetCurrentApplicationVersion()

        {

            Assembly asm = Assembly.GetExecutingAssembly();
            var fullName = asm.FullName;
            //"NIM-Addin, Version=1.0.1.0, Culture=neutral, PublicKeyToken=null"
            var version = fullName.Split(@", ".ToArray())[2].Split('=')[1];
            return version;

        }
        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("我们将读到EXCEL内容，并将内容提交到服务器，以生成证书");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("在这里输入登录信息");
        }


        private void InitPathManager()
        {
            InitRecourses.Init();
            var app = Globals.ThisAddIn.Application;

            Excel.Workbook oWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            oWorkbook.Save();


            var excelFileName = Globals.ThisAddIn.Application.ActiveWorkbook.FullName;

            var pathManager = IFilePathManagerProvider.PathProvider.TemplateFilesPath;
        }

        private void GenerateWord()
        {

            var app = Globals.ThisAddIn.Application;

            Excel.Workbook oWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            oWorkbook.Save();


            var excelFileName = Globals.ThisAddIn.Application.ActiveWorkbook.FullName;

            this.InitPathManager();

            var parser = NIM.CertificationGenerator.Core.WordParseFactory.GetWordParse(excelFileName);

            var wordResultFileName = parser.GeneraterFile();


            var docFileName = Path.Combine(Path.GetDirectoryName(wordResultFileName), Path.GetFileNameWithoutExtension(wordResultFileName) + ".doc");

            var document = new Aspose.Words.Document(wordResultFileName);
            try
            {
                document.Save(docFileName, Aspose.Words.SaveFormat.Doc);
            }
            catch (Exception ex)
            {
                this.ShowMessage($@"自动删除上次生成的证书文件失败,请手动删除它{docFileName},({ex.Message})", true);
                return;
            }
            System.IO.File.Delete(wordResultFileName);


            wordResultFileName = docFileName;
            this.ShowMessage($"WORD已生成,正在打开文档({wordResultFileName})...", false);


            var wordApp = new Word.Application();
            ((Word.Application)wordApp).Visible = true;

            wordApp.Documents.Open(wordResultFileName);
            this.ShowMessage(wordResultFileName, false);
            wordApp.Activate();

        }


        private void ShowMessage(string message, bool isError)
        {
            if (isError)
                this.richTextBox1.ForeColor = System.Drawing.Color.Red;
            else
                this.richTextBox1.ForeColor = System.Drawing.Color.Black;
            this.richTextBox1.Text = message;
            this.richTextBox1.Refresh();

        }
        private void ShowMessage(Exception ex)
        {
            while (ex.InnerException != null)
                ex = ex.InnerException;
            this.ShowMessage(ex.Message, true);

        }
        private void btnGenerator_Click(object sender, EventArgs e)
        {
            ////Get the assembly information
            //System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();

            ////Location is where the assembly is run from 
            //string assemblyLocation = assemblyInfo.Location;
            //MessageBox.Show(assemblyLocation);
            this.ShowMessage("", false);
            try
            {
                this.btnGenerator.Enabled = false;
                this.GenerateWord();
            }
            catch (Exception ex)
            {
                this.ShowMessage(ex);
            }
            finally
            {
                this.btnGenerator.Enabled = true;
            }
        }


   
    }
}

