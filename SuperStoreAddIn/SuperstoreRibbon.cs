using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using oi = Microsoft.Office.Interop;

namespace SuperStoreAddIn
{
    public partial class SuperstoreRibbon
    {
        private void SuperstoreRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnHighlightHighSales_Click(object sender, RibbonControlEventArgs e)
        {
            oi.Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;
            oi.Excel.Range usedRange = sheet.UsedRange;

            int salesColumn = 0;
            for (int i = 1; i <= usedRange.Columns.Count; i++)
            {
                if ((usedRange.Cells[1, i] as oi.Excel.Range).Value2.ToString() == "Sales")
                {
                    salesColumn = i;
                    break;
                }
            }

            if (salesColumn == 0) return;

            for (int i = 2; i <= usedRange.Rows.Count; i++)
            {
                var cell = usedRange.Cells[i, salesColumn] as oi.Excel.Range;
                double value = Convert.ToDouble(cell.Value2);
                if (value > 500)
                {
                    (cell.EntireRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                }
            }
        }

        private void btnShowSummary_Click(object sender, RibbonControlEventArgs e)
        {
            oi.Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;
            oi.Excel.Range usedRange = sheet.UsedRange;

            double totalSales = 0;
            int salesColumn = 0;

            for (int i = 1; i <= usedRange.Columns.Count; i++)
            {
                if ((usedRange.Cells[1, i] as oi.Excel.Range).Value2.ToString() == "Sales")
                {
                    salesColumn = i;
                    break;
                }
            }

            if (salesColumn == 0) return;

            for (int i = 2; i <= usedRange.Rows.Count; i++)
            {
                var cell = usedRange.Cells[i, salesColumn] as oi.Excel.Range;
                totalSales += Convert.ToDouble(cell.Value2);
            }

            int orderCount = usedRange.Rows.Count - 1;
            MessageBox.Show($"Total Sales: ${totalSales:N2}\nTotal Orders: {orderCount}", "Sales Summary");
        }

        private void btnClearHighlights_Click(object sender, RibbonControlEventArgs e)
        {
            oi.Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;
            oi.Excel.Range usedRange = sheet.UsedRange;
            usedRange.Interior.ColorIndex = oi.Excel.XlColorIndex.xlColorIndexNone;
        }
    }
}
