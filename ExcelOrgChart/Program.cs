using ExcelOrgChart.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelOrgChart
{
    class Program
    {
        private static Excel.Workbook Wb = null;
        private static Excel.Application Xl = null;
        private static Excel.Worksheet Sheet = null;
        private static Excel.Worksheet LinkSheet = null;

        static void Main(string[] args)
        {
            Xl = new Excel.Application();
            Xl.Visible = true;
            Wb = Xl.Workbooks.Add();
            LinkSheet = Wb.Worksheets[1];
            LinkSheet.Name = "LinkedSheet";

            Sheet = Wb.Worksheets.Add();
            Sheet.Name = "OrgChart";


            var myLayout = Xl.SmartArtLayouts[93];

            var smartArtShape = Sheet.Shapes.AddSmartArt(myLayout, 50, 50, 600, 600);

            smartArtShape.AlternativeText = "Test";

            if (smartArtShape.HasSmartArt == Office.MsoTriState.msoTrue)
            {
                Office.SmartArt smartArt = smartArtShape.SmartArt;
                Office.SmartArtNodes nds = smartArt.AllNodes;

                //Delete template nodes
                for (int i = nds.Count; i >= 1; i--)
                {
                    nds[i].Delete();
                }

                //Add main node
                Office.SmartArtNode main = smartArt.Nodes.Add();
                main.TextFrame2.TextRange.Text = "Node 1";
                var t = Sheet.Cells[1, 1];

                //Sheet.Hyperlinks.Add(main.TextFrame2, "", "'" + LinkSheet.Name + "'!A1", "", "");

                //Add main child node
                Office.SmartArtNode aNode = main.AddNode(Office.MsoSmartArtNodePosition.msoSmartArtNodeBelow);
                aNode.TextFrame2.TextRange.Text = "Node 1.1";
                //Add 1.1 child node
                Office.SmartArtNode a2Node = aNode.AddNode(Office.MsoSmartArtNodePosition.msoSmartArtNodeBelow);
                a2Node.TextFrame2.TextRange.Text = "Node 1.1.1";

                //Add main child node
                Office.SmartArtNode bNode = main.AddNode(Office.MsoSmartArtNodePosition.msoSmartArtNodeBelow);
                bNode.TextFrame2.TextRange.Text = "Node 1.2";

                //Add main child node
                Office.SmartArtNode cNode = main.AddNode(Office.MsoSmartArtNodePosition.msoSmartArtNodeBelow);
                cNode.TextFrame2.TextRange.Text = "Node 1.3";

                //Add main child node
                Office.SmartArtNode dNode = main.AddNode(Office.MsoSmartArtNodePosition.msoSmartArtNodeBelow);
                dNode.TextFrame2.TextRange.Text = "Node 1.4";
            }
        }

        private Node GetMainNode()
        {
            return new Node
            {
                Id = 1,
                Title = "Main",
                Parent = null,
                ParentID = 0,
                Children = new List<Node>
               {
                   new Node
                   {
                       Id = 2,
                       Title = "Sub 1",
                       ParentID = 1,
                       Children = null
                   },
                   new Node
                   {
                       Id = 3,
                       Title = "Sub 2",
                       ParentID = 1,
                       Children = null
                   },
                   new Node
                   {
                       Id = 4,
                       Title = "Sub 3",
                       ParentID = 1,
                       Children = new List<Node>
                       {
                           new Node
                           {
                               Id = 6,
                               Title = "Sub 3.1",
                               ParentID = 1,
                               Children = null
                           }
                       }
                   },
                   new Node
                   {
                       Id = 5,
                       Title = "Sub 4",
                       ParentID = 1,
                       Children = null
                   }
               }
            };
        }
    }
}
