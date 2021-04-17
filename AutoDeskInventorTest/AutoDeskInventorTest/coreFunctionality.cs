using System;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using Inventor;
using System.Windows;

namespace AutoDeskInventorTest
{
    class coreFunctionality
    {
        public static void toolTest(string inputExcelData, string template)
        {
            Inventor.Application _invApp;
            try
            {
                _invApp = (Inventor.Application)Marshal.GetActiveObject("Inventor.Application");
                createDrawing(_invApp, inputExcelData, template);
            }
            catch (Exception ex)
            {
                Console.WriteLine("error occured: " + ex.Message);
                try
                {
                    Type invAppType = Type.GetTypeFromProgID("Inventor.Application");
                    _invApp = (Inventor.Application)Activator.CreateInstance(invAppType);
                    _invApp.Visible = true;
                    createDrawing(_invApp, inputExcelData, template);
                }
                catch (Exception ex2)
                {
                    Console.WriteLine(ex2.ToString());
                    Console.WriteLine("Unable to get or start Inventor");
                }
            }
        }

        private static void createDrawing(Inventor.Application _invApp, string inputExcelData, string template)
        {
            try
            {
                Console.WriteLine("Inventor captured");
                System.Threading.Thread.Sleep(5000);
                DrawingDocument idwDoc = _invApp.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject, template, true) as DrawingDocument;
                Sheet sheet = idwDoc.ActiveSheet;
                DrawingSketch sketch = null;
                TransientGeometry tg = _invApp.TransientGeometry;
                if (sheet.Sketches.Count > 0)
                    sketch = sheet.Sketches[1];
                else
                    sketch = sheet.Sketches.Add();
                DataTable inputData = GetDataTableFromExcel(inputExcelData, true);
                bool fileCreated = false;
                foreach (DataRow iRow in inputData.Rows)
                {
                    if (iRow.ItemArray[0].ToString().ToLower() == "rectangle")
                    {
                        fileCreated = CreateRectangle(iRow, sheet, sketch, tg);
                    }
                    else if (iRow.ItemArray[0].ToString().ToLower() == "circle")
                    {
                        fileCreated = CreateCircle(iRow, sheet, sketch, tg);
                    }
                }
                Console.WriteLine(sheet.Height);
                Console.WriteLine(sheet.Width);
                string dir = System.IO.Path.GetDirectoryName(inputExcelData);
                string resultPath = System.IO.Path.Combine(dir, "ToolTestResult.dwg");
                idwDoc.SaveAsInventorDWG(resultPath, true);
                MessageBox.Show("Drawing successfully created!");
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
        private static bool CreateRectangle(DataRow iRow, Sheet iSheet, DrawingSketch iSketch, TransientGeometry tg)
        {
            bool isRectangleCreated = false;
            try
            {
                string tempCoord = iRow.ItemArray[4].ToString();
                string[] coord = tempCoord.Split(',');
                Double initX = Double.Parse(coord[0]);
                Double initY = Double.Parse(coord[1]);
                Double height = Double.Parse(iRow.ItemArray[1].ToString());
                Double width = Double.Parse(iRow.ItemArray[2].ToString());
                iSketch.Edit();
                Point2d pt1 = tg.CreatePoint2d(initX, initY);
                Point2d pt2 = tg.CreatePoint2d(initX + width, initY);
                Point2d pt3 = tg.CreatePoint2d(initX + width, initY + height);
                Point2d pt4 = tg.CreatePoint2d(initX, initY + height);
                SketchLine l1 = iSketch.SketchLines.AddByTwoPoints(pt1, pt2);
                SketchLine l2 = iSketch.SketchLines.AddByTwoPoints(pt2, pt3);
                iSketch.SketchLines.AddByTwoPoints(pt3, pt4);
                iSketch.SketchLines.AddByTwoPoints(pt4, pt1);
                iSketch.ExitEdit();
                GeometryIntent oGeo1 = iSheet.CreateGeometryIntent(l1, null);
                GeometryIntent oGeo2 = iSheet.CreateGeometryIntent(l2, null);
                LinearGeneralDimension iDim = iSheet.DrawingDimensions.GeneralDimensions.AddLinear(pt1, oGeo1);
                DimensionStyle iStyle = iDim.Style;
                iStyle.PartOffset = 45.0;
                iStyle.ShowDimensionLine = true;
                iSheet.DrawingDimensions.GeneralDimensions.AddLinear(pt2, oGeo2).Style = iStyle;
                isRectangleCreated = true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return isRectangleCreated;
        }

        private static bool CreateCircle(DataRow iRow, Sheet iSheet, DrawingSketch iSketch, TransientGeometry tg)
        {
            bool isCircleCreated = false;
            try
            {
                int circleInst = Int32.Parse(iRow.ItemArray[iRow.ItemArray.Count() - 1].ToString());
                int startNum = 0;
                string tempCoord = iRow.ItemArray[4].ToString();
                string[] coord = tempCoord.Split(',');
                Double initX = Double.Parse(coord[0]);
                Double initY = Double.Parse(coord[1]);
                Double hOffset = Double.Parse(iRow.ItemArray[5].ToString());
                Double vOffset = Double.Parse(iRow.ItemArray[6].ToString());
                Double radius = Double.Parse(iRow.ItemArray[3].ToString()) / 2;
                do
                {
                    iSketch.Edit();
                    Point2d cen;
                    if (startNum < circleInst / 2)
                        cen = tg.CreatePoint2d(initX + (hOffset * startNum), initY);
                    else
                        cen = tg.CreatePoint2d(initX + (hOffset * (startNum - 4)), initY + vOffset);
                    SketchCircle iCircle = iSketch.SketchCircles.AddByCenterRadius(cen, 2);
                    GeometryIntent oGeo1 = iSheet.CreateGeometryIntent(iCircle, null);
                    iSketch.ExitEdit();
                    iSheet.DrawingDimensions.GeneralDimensions.AddDiameter(cen, oGeo1, false, false, false);
                    isCircleCreated = true;
                    startNum += 1;
                } while (startNum < circleInst);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return isCircleCreated;
        }

        public static DataTable GetDataTableFromExcel(string path, bool hasHeader = true)
        {
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = System.IO.File.OpenRead(path))
                {
                    pck.Load(stream);
                }

                var ws = pck.Workbook.Worksheets.First();
                foreach (var worksheet in pck.Workbook.Worksheets)
                {
                    if (worksheet.Name == "Qty in Checkedin1")
                        ws = worksheet;
                }

                DataTable tbl = new DataTable();
                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                {
                    tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                }
                var startRow = hasHeader ? 2 : 1;
                for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                    DataRow row = tbl.Rows.Add();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                }
                return tbl;
            }
        }
    }
}
