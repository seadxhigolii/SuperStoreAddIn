using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using oi = Microsoft.Office.Interop;

namespace SuperStoreAddIn
{
    public partial class SuperstoreRibbon
    {
        private void SuperstoreRibbon_Load(object sender, RibbonUIEventArgs e) { }

        private void OptimizeExcelAction(Action<oi.Excel.Worksheet, object[,], oi.Excel.Range> action)
        {
            var app = Globals.ThisAddIn.Application;
            var sheet = app.ActiveSheet as oi.Excel.Worksheet;
            var usedRange = sheet.UsedRange;
            object[,] values = usedRange.Value2 as object[,];

            app.ScreenUpdating = false;
            app.EnableEvents = false;
            app.Calculation = oi.Excel.XlCalculation.xlCalculationManual;

            try
            {
                action(sheet, values, usedRange);
            }
            finally
            {
                app.ScreenUpdating = true;
                app.EnableEvents = true;
                app.Calculation = oi.Excel.XlCalculation.xlCalculationAutomatic;
            }
        }

        private void btnHighlightHighSales_Click(object sender, RibbonControlEventArgs e)
        {
            OptimizeExcelAction((sheet, values, usedRange) =>
            {
                int salesColumn = GetColumnIndex(values, "Sales");
                if (salesColumn == 0) return;

                for (int i = 2; i <= values.GetLength(0); i++)
                {
                    if (values[i, salesColumn] is double value && value > 500)
                    {
                        var cell = usedRange.Cells[i, salesColumn] as oi.Excel.Range;
                        cell.EntireRow.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                    }
                }
            });
        }

        private void btnShowSummary_Click(object sender, RibbonControlEventArgs e)
        {
            OptimizeExcelAction((sheet, values, usedRange) =>
            {
                int salesColumn = GetColumnIndex(values, "Sales");
                if (salesColumn == 0) return;

                double totalSales = 0;
                int orderCount = values.GetLength(0) - 1;

                for (int i = 2; i <= values.GetLength(0); i++)
                {
                    if (values[i, salesColumn] is double val) totalSales += val;
                }

                MessageBox.Show($"Total Sales: ${totalSales:N2}\nTotal Orders: {orderCount}", "Sales Summary");
            });
        }

        private void btnClearHighlights_Click(object sender, RibbonControlEventArgs e)
        {
            var sheet = Globals.ThisAddIn.Application.ActiveSheet as oi.Excel.Worksheet;
            var usedRange = sheet.UsedRange;
            usedRange.Interior.ColorIndex = oi.Excel.XlColorIndex.xlColorIndexNone;
        }

        private void btnRegionSales_Click(object sender, RibbonControlEventArgs e)
        {
            OptimizeExcelAction((sheet, values, usedRange) =>
            {
                int regionCol = GetColumnIndex(values, "Region");
                int salesCol = GetColumnIndex(values, "Sales");
                if (regionCol == 0 || salesCol == 0)
                {
                    MessageBox.Show("Could not find 'Region' or 'Sales' columns.");
                    return;
                }

                var regionSales = new Dictionary<string, double>();

                for (int row = 2; row <= values.GetLength(0); row++)
                {
                    var region = values[row, regionCol]?.ToString();
                    if (!string.IsNullOrEmpty(region) && values[row, salesCol] is double sale)
                    {
                        if (regionSales.ContainsKey(region)) regionSales[region] += sale;
                        else regionSales[region] = sale;
                    }
                }

                var summarySheet = Globals.ThisAddIn.Application.Worksheets.Add();
                summarySheet.Name = "Region Sales Summary";
                summarySheet.Cells[1, 1] = "Region";
                summarySheet.Cells[1, 2] = "Total Sales";

                int summaryRow = 2;
                foreach (var entry in regionSales.OrderByDescending(x => x.Value))
                {
                    summarySheet.Cells[summaryRow, 1] = entry.Key;
                    summarySheet.Cells[summaryRow, 2] = entry.Value;
                    summaryRow++;
                }

                var dataRange = summarySheet.Range["A1", $"B{summaryRow - 1}"];
                dataRange.Font.Bold = true;
                dataRange.Columns.AutoFit();

                MessageBox.Show("Region Sales summary created successfully!");
            });
        }

        private void btnLateShipping_Click(object sender, RibbonControlEventArgs e)
        {
            OptimizeExcelAction((sheet, values, usedRange) =>
            {
                int orderDateCol = GetColumnIndex(values, "Order Date");
                int shipDateCol = GetColumnIndex(values, "Ship Date");
                if (orderDateCol == 0 || shipDateCol == 0)
                {
                    MessageBox.Show("Could not find 'Order Date' or 'Ship Date' columns.");
                    return;
                }

                int delayCol = usedRange.Columns.Count + 1;
                sheet.Cells[1, delayCol] = "Delay (Days)";
                int delayedCount = 0;

                for (int row = 2; row <= values.GetLength(0); row++)
                {
                    if (values[row, orderDateCol] is double o && values[row, shipDateCol] is double s)
                    {
                        int delay = (DateTime.FromOADate(s) - DateTime.FromOADate(o)).Days;
                        sheet.Cells[row, delayCol] = delay;
                        if (delay > 6)
                        {
                            usedRange.Rows[row].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSalmon);
                            delayedCount++;
                        }
                    }
                }

                (sheet.Columns[delayCol] as oi.Excel.Range).AutoFit();
                MessageBox.Show($"Highlighted {delayedCount} late shipments (delay > 6 days).", "Late Shipping");
            });
        }

        private void btnCategorySummary_Click(object sender, RibbonControlEventArgs e)
        {
            OptimizeExcelAction((sheet, values, usedRange) =>
            {
                int categoryCol = GetColumnIndex(values, "Category");
                int salesCol = GetColumnIndex(values, "Sales");
                if (categoryCol == 0 || salesCol == 0)
                {
                    MessageBox.Show("Could not find 'Category' or 'Sales' columns.");
                    return;
                }

                var summary = new Dictionary<string, (double Total, int Count)>();

                for (int row = 2; row <= values.GetLength(0); row++)
                {
                    var cat = values[row, categoryCol]?.ToString();
                    if (!string.IsNullOrEmpty(cat) && values[row, salesCol] is double sales)
                    {
                        if (summary.ContainsKey(cat))
                            summary[cat] = (summary[cat].Total + sales, summary[cat].Count + 1);
                        else
                            summary[cat] = (sales, 1);
                    }
                }

                var summarySheet = Globals.ThisAddIn.Application.Worksheets.Add();
                summarySheet.Name = "Category Summary";
                summarySheet.Cells[1, 1] = "Category";
                summarySheet.Cells[1, 2] = "Total Sales";
                summarySheet.Cells[1, 3] = "Average Sale";
                summarySheet.Cells[1, 4] = "Order Count";

                int rowIdx = 2;
                foreach (var entry in summary)
                {
                    summarySheet.Cells[rowIdx, 1] = entry.Key;
                    summarySheet.Cells[rowIdx, 2] = entry.Value.Total;
                    summarySheet.Cells[rowIdx, 3] = entry.Value.Total / entry.Value.Count;
                    summarySheet.Cells[rowIdx, 4] = entry.Value.Count;
                    rowIdx++;
                }

                var dataRange = summarySheet.Range["A1", $"D{rowIdx - 1}"];
                dataRange.Columns.AutoFit();
                dataRange.Font.Bold = true;

                MessageBox.Show("Category summary created successfully!");
            });
        }

        private void btnSalesBuckets_Click(object sender, RibbonControlEventArgs e)
        {
            OptimizeExcelAction((sheet, values, usedRange) =>
            {
                int salesCol = GetColumnIndex(values, "Sales");
                if (salesCol == 0)
                {
                    MessageBox.Show("Could not find 'Sales' column.");
                    return;
                }

                int bucketCol = usedRange.Columns.Count + 1;
                sheet.Cells[1, bucketCol] = "Sales Bucket";

                for (int row = 2; row <= values.GetLength(0); row++)
                {
                    if (values[row, salesCol] is double sales)
                    {
                        string bucket;
                        var color = System.Drawing.Color.White;

                        if (sales < 100)
                        {
                            bucket = "Low";
                            color = System.Drawing.Color.LightGray;
                        }
                        else if (sales <= 500)
                        {
                            bucket = "Medium";
                            color = System.Drawing.Color.LightBlue;
                        }
                        else
                        {
                            bucket = "High";
                            color = System.Drawing.Color.LightGreen;
                        }

                        sheet.Cells[row, bucketCol] = bucket;
                        usedRange.Rows[row].Interior.Color = System.Drawing.ColorTranslator.ToOle(color);
                    }
                }

                (sheet.Columns[bucketCol] as oi.Excel.Range).AutoFit();
                MessageBox.Show("Sales Buckets assigned successfully!");
            });
        }

        private int GetColumnIndex(object[,] values, string columnName)
        {
            for (int col = 1; col <= values.GetLength(1); col++)
            {
                if (values[1, col]?.ToString() == columnName)
                    return col;
            }
            return 0;
        }
    }
}