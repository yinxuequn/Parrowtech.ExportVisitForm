using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace ExportVisitForm
{
    public partial class ExportVisitReportForm : Form
    {
        private ExcelHelper _excelHelper;

        public ExportVisitReportForm()
        {
            InitializeComponent();
        }

        private void ImportExcelButton_Click(object sender, EventArgs e)
        {
            try
            {
                string fileName = OpenFile();

                if (string.IsNullOrEmpty(fileName))
                {
                    OutputRichTextBox("没有选取文件，或者文件不存在.");
                    return;
                }

                _excelHelper = new ExcelHelper();
                OutputRichTextBox("开始导入Excel数据......");
                var soucesDataTable = ReadExcelToDataTable(fileName);
                OutputRichTextBox(string.Format("读取Excel数据 {0} 条.", soucesDataTable.Rows.Count));
                OutputRichTextBox("开始处理数据......");
                var soucesVisitRecords = CreateVisitRecords(soucesDataTable);
                OutputRichTextBox("开始统计数据......");
                var resultReportData = StatisticalData(soucesVisitRecords);

                if (resultReportData.Count == 0)
                {
                    OutputRichTextBox("需要导出的门店数量为0,请检查数据格式.");
                    return;
                }

                OutputRichTextBox("数据统计完成,开始导出到Excel.");
                string outExcelFileName = SaveFile();
                _excelHelper.VisitReportToExcel(resultReportData, outExcelFileName, "门店-总合");
                OutputRichTextBox("完成.");
            }
            catch (Exception ex)
            {
                OutputRichTextBox(ex.Message);
                throw;
            }
        }

        private string OpenFile()
        {
            string fileName = null;

            openFileDialog.Filter = "Excel 97-2003文档(*.xls)|*.xlsx|Excel 2007文档(*.xlsx)|*.xlsx";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                fileName = openFileDialog.FileName;
                OutputRichTextBox(string.Format("读取文件:{0}.", fileName));
            }

            return fileName;
        }

        private string SaveFile()
        {
            string fileName = null;
            saveFileDialog.Filter = "Excel 97-2003文档(*.xls)|*.xlsx|Excel 2007文档(*.xlsx)|*.xlsx";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                fileName = saveFileDialog.FileName;
                OutputRichTextBox(string.Format("保存到文件:{0}.", fileName));
            }

            return fileName;
        }

        private void OutputRichTextBox(string msg)
        {
            messageRichTextBox.AppendText(msg);
            messageRichTextBox.AppendText(Environment.NewLine);
        }

        /// <summary>
        /// 统计数据
        /// </summary>
        /// <param name="visitRecords"></param>
        private List<VisitReport> StatisticalData(List<VisitRecord> visitRecords)
        {
            var staticResult = visitRecords.GroupBy(a => a.StoreCode).Select(g =>
                (new { StoreCode = g.Key.ToString(), Count = g.Count() }));

            List<VisitReport> VisitReportResults = new List<VisitReport>();

            OutputRichTextBox(string.Format("需要处理 {0} 个门店.", staticResult.Count()));
            foreach (var result in staticResult)
            {
                if (result.Count < 2)
                {
                    continue;
                }

                var report = CreateVisitReport(result.StoreCode, result.Count, visitRecords);

                if (report != null)
                {
                    VisitReportResults.Add(report);
                }
            }

            return VisitReportResults;
        }

        private VisitReport CreateVisitReport(string storeCode, int count, List<VisitRecord> visitRecords)
        {
            var report = new VisitReport();

            report.StoreCode = storeCode;
            report.IrregularCount = count;

            var result = visitRecords.FindAll(s => s.StoreCode.Trim().ToUpper() == storeCode.Trim().ToUpper()).OrderBy(s => s.Date);

            if (result.Count() == 0)
            {
                return null;
            }

            report.VisiterCode = result.First().VisiterCode;
            report.VisiterName = result.First().VisiterName;
            report.StoreName = result.First().StoreName;
            report.StoreAddress = result.First().StoreAddress;
            report.VisitRecords = result.ToList<VisitRecord>();

            foreach (var item in result)
            {
                if (item.Deviation.Equals("基准点不准"))
                {
                    report.ReferencePointErrorCount++;
                }
                else if (item.Deviation.Equals("拜访点不准"))
                {
                    report.VisitPointErrorCount++;
                }
            }

            return report;
        }

        /// <summary>
        /// 读取Excel中的数据
        /// </summary>
        /// <returns></returns>
        private DataTable ReadExcelToDataTable(string filePath)
        {
            return _excelHelper.ExcelToDataTable(filePath, "Rawdata");
        }


        /// <summary>
        /// 将DataTable中的数据转换到实体中
        /// </summary>
        /// <param name="soucesDataTable"></param>
        /// <returns></returns>
        private List<VisitRecord> CreateVisitRecords(DataTable soucesDataTable)
        {
            List<VisitRecord> visitRecords = new List<VisitRecord>();
            OutputRichTextBox("数据验证......");
            int excelRow = 1;
            int successCount = 0;
            int errorCount = 0;
            string msg = null;

            foreach (DataRow dataRow in soucesDataTable.Rows)
            {
                var record = CreateVisitRecord(dataRow, ref msg);

                if (record != null)
                {
                    visitRecords.Add(record);
                    successCount++;
                }
                else
                {
                    OutputRichTextBox(string.Format("数据验证错误:Excel第 {0} 行数据的{1}", excelRow, msg));
                    errorCount++;
                }

                excelRow++;
            }

            OutputRichTextBox(string.Format("数据验证结果: 成功 {0} 条  失败 {1} 条.", successCount, errorCount));
            return visitRecords;
        }

        /// <summary>
        /// 将DataTable中的数据转换到实体中
        /// </summary>
        /// <param name="dataRow"></param>
        /// <returns></returns>
        private VisitRecord CreateVisitRecord(DataRow dataRow, ref string msg)
        {
            VisitRecord record = new VisitRecord();

            record.ID = Guid.NewGuid();

            if (string.IsNullOrEmpty(dataRow["拜访人编码"].ToString()))
            {
                msg = " 拜访人编码 为空";
                return null;
            }
            else
            {
                record.VisiterCode = dataRow["拜访人编码"].ToString().Trim();
            }

            record.VisiterName = dataRow["拜访人"].ToString().Trim(); ;

            if (string.IsNullOrEmpty(dataRow["门店编码"].ToString()))
            {
                msg = " 门店编码 为空";
                return null;
            }
            else
            {
                record.StoreCode = dataRow["门店编码"].ToString().Trim();
            }

            if (string.IsNullOrEmpty(dataRow["门店名称"].ToString()))
            {
                msg = " 门店名称 为空";
                return null;
            }
            else
            {
                record.StoreName = dataRow["门店名称"].ToString().Trim();
            }

            record.StoreAddress = dataRow["门店地址"].ToString().Trim();

            if (string.IsNullOrEmpty(dataRow["偏差情况"].ToString()))
            {
                msg = " 偏差情况 为空";
                return null;
            }
            else if (dataRow["偏差情况"].ToString().Trim() != "拜访点不准" && dataRow["偏差情况"].ToString().Trim() != "基准点不准")
            {
                msg = " 偏差情况 不是 拜访点不准 或者 基准点不准";
                return null;
            }
            else
            {
                record.Deviation = dataRow["偏差情况"].ToString().Trim();
            }

            record.DeviationReason = dataRow["产生偏差原因"].ToString().Trim();

            if (!string.IsNullOrEmpty(dataRow["拜访日期"].ToString()))
            {
                DateTime date;
                if (DateTime.TryParse(dataRow["拜访日期"].ToString(), out date))
                {
                    record.Date = date;
                }
                else
                {
                    msg = " 拜访日期 格式不正确";
                    return null;
                }
            }
            else
            {
                msg = " 拜访日期 为空";
                return null;
            }

            return record;
        }
    }
}
