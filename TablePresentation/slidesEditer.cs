using System;
using System.Data;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointOperator
{
    public class SlidesEditer
    {
        PowerPoint.Application appPPT = null;
        PowerPoint.Presentation pptPrest = null;

        public static void openPPT(string filePath)
        {
            PowerPoint.Application appPPT = new PowerPoint.Application();
            appPPT.Visible = MsoTriState.msoCTrue;
            PowerPoint.Presentation pptPrest = appPPT.Presentations.Open(filePath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoCTrue);
        }

        public static void addSilde(PowerPoint.Presentation pptPrest, int pageIndex, string title, DataRow titleRow, int game)
        {
            pptPrest.Slides[game].Copy();
            pptPrest.Slides.Paste(pageIndex);
            addRow(pptPrest, pageIndex, titleRow, 1);
        }

        public static void addRow(PowerPoint.Presentation pptPrest, int pageIndex, DataRow dataRow, int rowIndex = -1)
        {

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
