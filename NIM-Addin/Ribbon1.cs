using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace NIM_Addin
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

          
            this.toggleButton1.Checked = false;
        }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            InitRecourses.Init();
            var taskpane = TaskPaneManager.GetTaskPane("NIMConfig", "NIM Center", () => new UserControl1());
            taskpane.Visible = !taskpane.Visible;

            //var isVisibile = false;
            //if (Globals.ThisAddIn.CustomTaskPanes.Count == 0)
            //{
            //    Globals.ThisAddIn.AddCommandBar();
            //    isVisibile = true;
            //}
            //else
            //    isVisibile = this.toggleButton1.Checked;

            //Globals.ThisAddIn.CustomTaskPanes[0].Visible = isVisibile;



        }
    }
}
