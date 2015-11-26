using System;
using System.Data;
using System.IO;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointOperator
{
    public class SlidesEditer
    {
        PowerPoint.Application appPPT = null;
        PowerPoint.Presentation pptPrest = null;

        public static PowerPoint.Presentation openPPT(string filePath)
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine("Please put the PowerPoint templet file in the root directory.");
                Console.ReadKey();
                //return null;
                Environment.Exit(0);
            }
            PowerPoint.Application appPPT = new PowerPoint.Application();
            appPPT.Visible = MsoTriState.msoCTrue;
            PowerPoint.Presentation pptPrest = appPPT.Presentations.Open(filePath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoCTrue);
            return pptPrest;
        }

        public static void addSilde(PowerPoint.Presentation pptPrest, int pageIndex, string title, DataRow titleContent, int game)
        {
            pptPrest.Slides[game].Copy();
            pptPrest.Slides.Paste(pageIndex);
            addContent(pptPrest, pageIndex, 1, titleContent);
        }

        public static void addRow(PowerPoint.Presentation pptPrest, int pageIndex, DataRow rowContent)
        {
            pptPrest.Slides[pageIndex].Shapes[1].Table.Rows.Add(-1);
            addContent(pptPrest, pageIndex, -1, rowContent);
        }

        public static void addContent(PowerPoint.Presentation pptPrest, int pageIndex, int rowIndex, DataRow content)
        {
            int columnCount = pptPrest.Slides[pageIndex].Shapes[1].Table.Columns.Count;
            for(int i = 0; i < columnCount; i++)
            {
                pptPrest.Slides[pageIndex].Shapes[1].Table.Cell(rowIndex, i).Shape.TextFrame.TextRange.Text= ((dynamic)content[i])["text"];
                pptPrest.Slides[pageIndex].Shapes[1].Table.Cell(rowIndex, i).Shape.TextFrame.TextRange.Font.Color.RGB = ((dynamic)content[i])["color"]; //RGB(0, 0, 255)
                //pptPrest.Slides[pageIndex].Shapes[1].Table.Cell(rowIndex, i).Shape. to be continued
            }
        }

        private static void Dispose(string fileName, ref PowerPoint.Application appPPT, ref PowerPoint.Presentation pptPrest)
        {
            pptPrest.Close();
            appPPT.Quit();

            ReleaseObject(pptPrest);
            ReleaseObject(appPPT);
        }

        private static void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
