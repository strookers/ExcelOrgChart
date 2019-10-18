using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ConsoleApp1
{
    class Program
    {
        private static Excel.Workbook Wb = null;
        private static Excel.Application Xl = null;
        private static Excel.Worksheet Sheet = null;

        static void Main(string[] args)
        {
            Xl = new Excel.Application();
            Xl.Visible = true;
            Wb = Xl.Workbooks.Add();
            Sheet = Wb.Worksheets[1];

            var myLayout = Xl.SmartArtLayouts[88];

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

                //Add main child node
                Office.SmartArtNode aNode = main.AddNode();
                aNode.TextFrame2.TextRange.Text = "Node 1.1";
                //Add 1.1 child node
                Office.SmartArtNode a2Node = aNode.Nodes.Add();
                a2Node.TextFrame2.TextRange.Text = "Node 1.1.1";

                //Add main child node
                Office.SmartArtNode bNode = main.AddNode();
                bNode.TextFrame2.TextRange.Text = "Node 1.2";

                //Add main child node
                Office.SmartArtNode cNode = main.AddNode();
                cNode.TextFrame2.TextRange.Text = "Node 1.3";

                //Add main child node
                Office.SmartArtNode dNode = main.AddNode();
                dNode.TextFrame2.TextRange.Text = "Node 1.4";

            }
        }
    }
}
