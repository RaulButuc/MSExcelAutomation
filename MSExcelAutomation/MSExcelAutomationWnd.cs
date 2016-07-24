using System;
using System.Windows.Forms;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace MSExcelAutomation
{
    public partial class MSExcelAutomationWnd : Form
    {
        public MSExcelAutomationWnd()
        {
            InitializeComponent();
        }

        private void automateExcelSpreadsheet_Click(object sender, EventArgs e)
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;

            try
            {
                oXL = new Excel.Application();
                oXL.Visible = true;

                oWB = oXL.Workbooks.Add(Missing.Value);
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                oSheet.Cells[1, 1] = "First Name";
                oSheet.Cells[1, 2] = "Last Name";
                oSheet.Cells[1, 3] = "E-mail";
                oSheet.Cells[1, 4] = "Salary";

                oSheet.get_Range("A1", "D1").Font.Bold = true;
                oSheet.get_Range("A1", "D1").VerticalAlignment =
                    Excel.XlVAlign.xlVAlignCenter;

                string[,] saNames = new string[5, 2];

                saNames[0, 0] = "John";
                saNames[0, 1] = "Smith";
                saNames[1, 0] = "Tom";
                saNames[1, 1] = "Brown";
                saNames[2, 0] = "Sue";
                saNames[2, 1] = "Thomas";
                saNames[3, 0] = "Jane";
                saNames[3, 1] = "Jones";
                saNames[4, 0] = "Adams";
                saNames[4, 1] = "Johnson";

                oSheet.get_Range("A2", "B6").Value2 = saNames;

                oRng = oSheet.get_Range("C2", "C6");
                oRng.Formula = "=A2 & \" \" & B2";

                oRng = oSheet.get_Range("D2", "D6");
                oRng.Formula = "=RAND()*100000";
                oRng.NumberFormat = "$0.00";

                oRng = oSheet.get_Range("A1", "D1");
                oRng.EntireColumn.AutoFit();

                DisplayQuarterlySales(oSheet);

                oXL.Visible = true;
                oXL.UserControl = true;
            }
            catch (Exception exception)
            {
                string errorMessage = "Error :";

                errorMessage = string.Concat(errorMessage, exception.Message);
                errorMessage = string.Concat(errorMessage, " Line: ");
                errorMessage = string.Concat(errorMessage, exception.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }

        private void DisplayQuarterlySales(Excel._Worksheet oWS)
        {
            Excel._Workbook oWB;
            Excel.Series oSeries;
            Excel.Range oResizeRange;
            Excel._Chart oChart;
            string sMsg;
            int iNumQtrs;

            for (iNumQtrs = 4; iNumQtrs >=2; --iNumQtrs)
            {
                sMsg = "Enter sales data for ";
                sMsg = string.Concat(sMsg, iNumQtrs);
                sMsg = string.Concat(sMsg, " quarter(s)?");

                DialogResult iRet = MessageBox.Show(sMsg, "Quarterly Sales?",
                    MessageBoxButtons.YesNo);

                if (iRet == DialogResult.Yes)
                {
                    break;
                }
            }

            oResizeRange = oWS.get_Range("E1", "E1").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Formula = "=\"Q\" & COLUMN()-4 & CHAR(10) & \"Sales\"";

            oResizeRange.Orientation = 38;
            oResizeRange.WrapText = true;

            oResizeRange.Interior.ColorIndex = 36;

            oResizeRange = oWS.get_Range("E2", "E6").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Formula = "=RAND()*100";
            oResizeRange.NumberFormat = "$0.00";

            oResizeRange = oWS.get_Range("E1", "E6").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            oResizeRange = oWS.get_Range("E8", "E8").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Formula = "=SUM(E2:E6)";
            oResizeRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle =
                Excel.XlLineStyle.xlDouble;
            oResizeRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight =
                Excel.XlBorderWeight.xlThick;

            oWB = (Excel._Workbook)oWS.Parent;
            oChart = (Excel._Chart)oWB.Charts.Add(Missing.Value, Missing.Value,
                Missing.Value, Missing.Value);

            oResizeRange = oWS.get_Range("E2:E6", Missing.Value).
                get_Resize(Missing.Value, iNumQtrs);
            oChart.ChartWizard(oResizeRange, Excel.XlChartType.xl3DColumn, Missing.Value,
                Excel.XlRowCol.xlColumns, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            oSeries = (Excel.Series)oChart.SeriesCollection(1);
            oSeries.XValues = oWS.get_Range("A2", "A6");
            
            for (int iRet = 1; iRet <= iNumQtrs; ++iRet)
            {
                oSeries = (Excel.Series)oChart.SeriesCollection(iRet);
                string seriesName;
                seriesName = "=\"Q";
                seriesName = string.Concat(seriesName, iRet);
                seriesName = string.Concat(seriesName, "\"");
                oSeries.Name = seriesName;
            }

            oChart.Location(Excel.XlChartLocation.xlLocationAsObject, oWS.Name);

            oResizeRange = (Excel.Range)oWS.Rows.get_Item(10, Missing.Value);
            oWS.Shapes.Item("Chart 1").Top = (float)(double)oResizeRange.Top;
            oResizeRange = (Excel.Range)oWS.Columns.get_Item(2, Missing.Value);
            oWS.Shapes.Item("Chart 1").Left = (float)(double)oResizeRange.Left;
        }

        private void closeMainWnd_Click(object sender, EventArgs e)
        {
            ActiveForm.Close();
        }
    }
}
