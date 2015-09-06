using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.IO;
using System.Data;

namespace ExportVisitForm
{
    public class ExcelHelper : IDisposable
    {
        private IWorkbook workbook = null;
        private FileStream fs = null;
        private bool disposed;

        public ExcelHelper()
        {
            disposed = false;
        }

        /// <summary>
        /// 将DataTable数据导入到excel中
        /// </summary>
        /// <param name="data">要导入的数据</param>
        /// <param name="isColumnWritten">DataTable的列名是否要导入</param>
        /// <param name="sheetName">要导入的excel的sheet的名称</param>
        /// <returns>导入数据行数(包含列名那一行)</returns>
        public int VisitReportToExcel(List<VisitReport> reports, string reportFileName, string sheetName = null)
        {
            int currentRowIndex = 1;
            int maxVisitRecordsCount = 0;

            ISheet sheet = null;
            IRow columnRow = null;

            if (reportFileName.IndexOf(".xlsx") > 0) // 2007版本
                workbook = new XSSFWorkbook();
            else if (reportFileName.IndexOf(".xls") > 0) // 2003版本
                workbook = new HSSFWorkbook();

            if (workbook != null)
            {
                if (string.IsNullOrEmpty(sheetName))
                {
                    sheet = workbook.CreateSheet();
                }
                else
                {
                    sheet = workbook.CreateSheet(sheetName);
                }
            }
            else
            {
                return -1;
            }


            var firstRow = sheet.GetRow(0);
            if (firstRow == null)
            {
                columnRow = sheet.CreateRow(0);
            }
            else
            {
                columnRow = firstRow;
            }

            var column0 = columnRow.CreateCell(0);
            SetColumnStyle(workbook, column0, "拜访人编码");
            var column1 = columnRow.CreateCell(1);
            SetColumnStyle(workbook, column1, "拜访人");
            var column2 = columnRow.CreateCell(2);
            SetColumnStyle(workbook, column2, "门店编码");
            var column3 = columnRow.CreateCell(3);
            SetColumnStyle(workbook, column3, "门店名称");
            var column4 = columnRow.CreateCell(4);
            SetColumnStyle(workbook, column4, "门店地址");
            var column5 = columnRow.CreateCell(5);
            SetColumnStyle(workbook, column5, "不合规次数");
            var column6 = columnRow.CreateCell(6);
            SetColumnStyle(workbook, column6, "基准点不准次数");
            var column7 = columnRow.CreateCell(7);
            SetColumnStyle(workbook, column7, "拜访点不准次数");
            var column8 = columnRow.CreateCell(8);
            SetColumnStyle(workbook, column8, "多次基准点不准原因");

            sheet.SetColumnWidth(0, 10 * 256);
            sheet.SetColumnWidth(1, 8 * 256);
            sheet.SetColumnWidth(2, 10 * 256);
            sheet.SetColumnWidth(3, 30 * 256);
            sheet.SetColumnWidth(4, 20 * 256);
            sheet.SetColumnWidth(5, 10 * 256);
            sheet.SetColumnWidth(6, 10 * 256);
            sheet.SetColumnWidth(7, 10 * 256);
            sheet.SetColumnWidth(8, 10 * 256);

            var cellStyle = CreateCellStyle(workbook);

            foreach (var report in reports)
            {
                IRow row = sheet.CreateRow(currentRowIndex);
                var cell0 = row.CreateCell(0);
                cell0.CellStyle = cellStyle;
                cell0.SetCellValue(report.VisiterCode);

                var cell1 = row.CreateCell(1);
                cell1.CellStyle = cellStyle;
                cell1.SetCellValue(report.VisiterName);

                var cell2 = row.CreateCell(2);
                cell2.CellStyle = cellStyle;
                cell2.SetCellValue(report.StoreCode);

                var cell3 = row.CreateCell(3);
                cell3.CellStyle = cellStyle;
                cell3.SetCellValue(report.StoreName);

                var cell4 = row.CreateCell(4);
                cell4.CellStyle = cellStyle;
                cell4.SetCellValue(report.StoreAddress);

                var cell5 = row.CreateCell(5);
                cell5.CellStyle = cellStyle;
                cell5.SetCellValue(report.IrregularCount);

                var cell6 = row.CreateCell(6);
                cell6.CellStyle = cellStyle;
                cell6.SetCellValue(report.ReferencePointErrorCount);

                var cell7 = row.CreateCell(7);
                cell7.CellStyle = cellStyle;
                cell7.SetCellValue(report.VisitPointErrorCount);

                var cell8 = row.CreateCell(8);
                cell8.CellStyle = cellStyle;
                cell8.SetCellValue("");

                int currentCount = report.VisitRecords.Count();
                //动态创建列
                if (currentCount > maxVisitRecordsCount)
                {
                    for (int i = maxVisitRecordsCount; i < currentCount; i++)
                    {
                        sheet.SetColumnWidth(8 + i * 3 + 1, 10 * 256);
                        sheet.SetColumnWidth(8 + i * 3 + 2, 10 * 256);
                        sheet.SetColumnWidth(8 + i * 3 + 3, 10 * 256);

                        SetColumnStyle(workbook, columnRow.CreateCell(8 + i * 3 + 1), string.Format("第{0}次", i + 1));
                        SetColumnStyle(workbook, columnRow.CreateCell(8 + i * 3 + 2), "不合规原因");
                        SetColumnStyle(workbook, columnRow.CreateCell(8 + i * 3 + 3), "原因分析");
                    }

                    maxVisitRecordsCount = currentCount;
                }

                for (int i = 0; i < currentCount; i++)
                {
                    var cell10 = row.CreateCell(8 + i * 3 + 1);
                    cell10.CellStyle = cellStyle;
                    cell10.SetCellValue(report.VisitRecords[i].Date.ToShortDateString());

                    var cell11 = row.CreateCell(8 + i * 3 + 2);
                    cell11.CellStyle = cellStyle;
                    cell11.SetCellValue(report.VisitRecords[i].Deviation);

                    var cell12 = row.CreateCell(8 + i * 3 + 3);
                    cell12.CellStyle = cellStyle;
                    cell12.SetCellValue(report.VisitRecords[i].DeviationReason);
                }

                currentRowIndex++;
            }

            using (fs = new FileStream(reportFileName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                workbook.Write(fs);
            }

            return currentRowIndex;

        }

        private void SetColumnStyle(IWorkbook hssfworkbook, ICell cell, string name)
        {
            var cellStyle = hssfworkbook.CreateCellStyle();
            cellStyle.WrapText = true;
            cellStyle.Alignment = HorizontalAlignment.Center;
            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            IFont font = hssfworkbook.CreateFont();
            font.FontName = "宋体";
            font.FontHeightInPoints = 10;
            font.Boldweight = 10;
            cellStyle.SetFont(font);
            cell.CellStyle = cellStyle;
            cell.CellStyle.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.BlueGrey.Index;
            cell.SetCellValue(name);
        }

        private ICellStyle CreateCellStyle(IWorkbook hssfworkbook)
        {
            var cellStyle = hssfworkbook.CreateCellStyle();
            cellStyle.WrapText = true;
            cellStyle.Alignment = HorizontalAlignment.Left;
            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            IFont font = hssfworkbook.CreateFont();
            font.FontName = "宋体";
            font.FontHeightInPoints = 10;
            cellStyle.SetFont(font);

            return cellStyle;
        }
        /// <summary>
        /// 将excel中的数据导入到DataTable中
        /// </summary>
        /// <param name="sheetName">excel工作薄sheet的名称</param>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
        /// <returns>返回的DataTable</returns>
        public DataTable ExcelToDataTable(string fileName, string sheetName = null, bool isFirstRowColumn = true)
        {
            ISheet sheet = null;
            DataTable data = new DataTable();
            int startRow = 0;

            fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                workbook = new XSSFWorkbook(fs);
            else if (fileName.IndexOf(".xls") > 0) // 2003版本
                workbook = new HSSFWorkbook(fs);

            if (!string.IsNullOrEmpty(sheetName))
            {
                sheet = workbook.GetSheet(sheetName);
                if (sheet == null) //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                {
                    sheet = workbook.GetSheetAt(0);
                }
            }
            else
            {
                sheet = workbook.GetSheetAt(0);
            }
            if (sheet != null)
            {
                IRow firstRow = sheet.GetRow(0);
                int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数

                if (isFirstRowColumn)
                {
                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                    {
                        ICell cell = firstRow.GetCell(i);
                        if (cell != null)
                        {
                            string cellValue = cell.StringCellValue;
                            if (cellValue != null)
                            {
                                DataColumn column = new DataColumn(cellValue);
                                data.Columns.Add(column);
                            }
                        }
                    }
                    startRow = sheet.FirstRowNum + 1;
                }
                else
                {
                    startRow = sheet.FirstRowNum;
                }

                //最后一列的标号
                int rowCount = sheet.LastRowNum;
                for (int i = startRow; i <= rowCount; ++i)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue; //没有数据的行默认是null　　　　　　　

                    DataRow dataRow = data.NewRow();
                    for (int j = row.FirstCellNum; j < cellCount; ++j)
                    {
                        var cell = row.Cells[j];

                        if (cell != null) //同理，没有数据的单元格都默认是null
                        {

                            if (j == 0)
                            {
                                dataRow[j] = cell.DateCellValue;
                            }
                            else
                            {
                                dataRow[j] = cell.ToString();
                            }

                        }
                    }
                    data.Rows.Add(dataRow);
                }
            }

            return data;

        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                if (disposing)
                {
                    if (fs != null)
                        fs.Close();
                }

                fs = null;
                disposed = true;
            }
        }
    }
}
