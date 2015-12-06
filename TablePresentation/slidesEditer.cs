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
                //Environment.Exit(0);
                Program.showMenu();
            }
            PowerPoint.Application appPPT = new PowerPoint.Application();
            //appPPT.Visible = MsoTriState.msoFalse;
            PowerPoint.Presentation pptPrest = appPPT.Presentations.Open(filePath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoCTrue);
            return pptPrest;
        }

        public static void addSilde(PowerPoint.Presentation pptPrest, int pageIndex, string title, DataRow titleContent, DataRow firstRow, int game)
        {

            pptPrest.Slides[game + 2].Copy();
            pptPrest.Slides[game + 2].Copy();
            if (pageIndex < 0 | pageIndex > pptPrest.Slides.Count)
            {
                pageIndex = pptPrest.Slides.Count;
                pptPrest.Slides.Paste(pageIndex);
            }
            else
            {
                pptPrest.Slides.Paste(pageIndex);
            }  
            pptPrest.Slides[pageIndex].Shapes.Title.TextFrame.TextRange.Text = title;
            addContent(pptPrest, pageIndex, 1, titleContent);
            addContent(pptPrest, pageIndex, 2, firstRow);
        }

        public static void addRow(PowerPoint.Presentation pptPrest, int pageIndex, DataRow rowContent)
        {
            pptPrest.Slides[pageIndex].Shapes[1].Table.Rows.Add();
            int n =pptPrest.Slides[pageIndex].Shapes[1].Table.Rows.Count;
            addContent(pptPrest, pageIndex, n, rowContent);
        }

        public static void addContent(PowerPoint.Presentation pptPrest, int pageIndex, int rowIndex, DataRow content)
        {
            int columnCount = pptPrest.Slides[pageIndex].Shapes[1].Table.Columns.Count;
            for(int i = 0; i < columnCount; i++)
            {
                pptPrest.Slides[pageIndex].Shapes[1].Table.Cell(rowIndex, i+1).Shape.TextFrame.TextRange.Text= ((dynamic)content[i+1])["text"];
                //pptPrest.Slides[pageIndex].Shapes[1].Table.Cell(rowIndex, i).Shape.TextFrame.TextRange.Font.Color.RGB = ((dynamic)content[i])["color"]; //RGB(0, 0, 255)
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
