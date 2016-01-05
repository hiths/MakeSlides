using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;

namespace Addins
{
    public class Functions
    {
        public static int RGBToIntBGR(string argb)
        {
            Color c = ColorTranslator.FromHtml("#" + argb);
            return c.B * (int)Math.Pow(16, 4) + c.G * (int)Math.Pow(16, 2) + c.R;
        }

        public static void regulateData(DataTable dt, int width)
        {
            if (dt.Columns.Count > width)
            {
                for (int i = width; i < dt.Columns.Count; i++)
                {
                    dt.Columns.RemoveAt(i);
                }
            }
            else
            {
                width = dt.Columns.Count;
            }
            foreach (DataRow dr in dt.Rows)
            {
                for (int i = 0; i < width; i++)
                {
                    bool isColor = false;
                    if (!string.IsNullOrWhiteSpace(((dynamic)dr[i])["color"]))
                    {
                        int pptColor = Functions.RGBToIntBGR(((dynamic)dr[i])["color"]);
                        isColor = true;
                        dr[i] = new Dictionary<string, object> { { "text", ((dynamic)dr[i])["text"] }, { "color", ((dynamic)dr[i])["color"] }, { "pptColor", pptColor }, { "isColor", isColor } };
                    }
                    else
                    {
                        dr[i] = new Dictionary<string, object> { { "text", ((dynamic)dr[i])["text"] }, { "isColor", isColor } };
                    }
                }
            }
        }
    }
}