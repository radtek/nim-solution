using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NIM_Addin
{
    public class TaskPaneManager
    {
        static Dictionary<string, Microsoft.Office.Tools.CustomTaskPane> _createdPanes = new Dictionary<string, Microsoft.Office.Tools.CustomTaskPane>();

        /// <summary>
        /// Gets the taskpane by name (if exists for current excel window then returns existing instance, otherwise uses taskPaneCreatorFunc to create one). 
        /// </summary>
        /// <param name="taskPaneId">Some string to identify the taskpane</param>
        /// <param name="taskPaneTitle">Display title of the taskpane</param>
        /// <param name="taskPaneCreatorFunc">The function that will construct the taskpane if one does not already exist in the current Excel window.</param>
        public static Microsoft.Office.Tools.CustomTaskPane GetTaskPane(string taskPaneId, string taskPaneTitle, Func<UserControl> taskPaneCreatorFunc)
        {
            string key = string.Format("{0}({1})", taskPaneId, Globals.ThisAddIn.Application.Hwnd);
            if (!_createdPanes.ContainsKey(key))
            {
                //this.objPanel.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;
               
                var pane = Globals.ThisAddIn.CustomTaskPanes.Add(taskPaneCreatorFunc(), taskPaneTitle);
                pane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionLeft;
                _createdPanes[key] = pane;
            }
            return _createdPanes[key];
        }
    }
}
//https://stackoverflow.com/questions/34027785/excel-vsto-addin-showing-hiding-taskpane



/*
 * 
 * 
 * TableCellProperties tableCellProperties1 = new TableCellProperties();
        TableCellWidth tableCellWidth1 = new TableCellWidth(){ Width = "1795", Type = TableWidthUnitValues.Dxa };


tableCell1.Append(tableCellProperties1);



    var fileName = @"C:\Users\Exten\Documents\A.docx";
            using (WordprocessingDocument document
               = WordprocessingDocument.Open(fileName, true))
            {


                var table = document.MainDocumentPart.Document.Body.Descendants<Table>().Skip(1).First();
                var rows = table.Elements<TableRow>().ToList();
                var cellIndex = 2;
                for(var i = 0;i<rows.Count;i++)
                {
                    var row = rows[i];
                    var count = row.Elements<TableCell>().Count();

                    var cell = row.ElementAt(count);
                    row.RemoveChild(cell);

                }

                document.Save();



            }

*/

