using System;
using System.Data;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointOperator
{
    public class SlidesEditer
    {
        public static PowerPoint.Presentation openPPT(string filePath, PowerPoint.Application appPPT)
        {
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
            //pptPrest.SlideShowWindow.Presentation.Slides[pageIndex].Select();
            pptPrest.Slides[pageIndex].Select();
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
                if((((dynamic)content[i + 1])["color"]).GetType() == Type.GetType("System.Int32"))
                {
                    pptPrest.Slides[pageIndex].Shapes[1].Table.Cell(rowIndex, i + 1).Shape.TextFrame.TextRange.Font.Color.RGB = ((dynamic)content[i + 1])["color"];
                }
            }
        }
    }
}
