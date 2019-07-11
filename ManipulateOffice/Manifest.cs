using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using OfficeOpenXml;
using Spire.Xls;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;

//Todo:User interface, change format to xlsx, xls
namespace Manifest
{
    class ReportGenerate
    {
        public string manifestFileName { get; set; }
        public string templateFileName { get; set; }
        public string manifestFileDir { get; set; }
        public string templateFileDir { get; set; }

        public ExcelTypeConverter converter { get; set; }

        public ReportGenerate()
        {
            converter = new ExcelTypeConverter();
        }

        public void ExcelGenearte()
        {
            var temp = new FileInfo(templateFileDir + Path.DirectorySeparatorChar + templateFileName);
            var originManiName = manifestFileDir + Path.DirectorySeparatorChar + manifestFileName;
            FileInfo mani;

            if (manifestFileName.EndsWith(".xls"))
            {
                var originMani = new FileInfo(originManiName);
                mani = new FileInfo(originMani + "x");
                var IsConverted = converter.XlsToXlsx(originMani);
                if (!IsConverted)
                {
                    Console.WriteLine("Convert fail.");
                    return;
                }
            }
            else
            {
                mani = new FileInfo(originManiName);
            }

            using (ExcelPackage maniPackage = new ExcelPackage(mani))
            {
                var maniWs = maniPackage.Workbook.Worksheets.FirstOrDefault();
                //To Do clean up the Data
                using (ExcelPackage tempPackage = new ExcelPackage(temp))
                {
                    //增值税清单模板
                    var tempWs = tempPackage.Workbook.Worksheets.FirstOrDefault();
                    int currentIndex = 2;
                    int count = 1;
                    int preIndex = 2;
                    double moneySum = 0;
                    string outputDir = @"F:\清单结果";
                    string fileName = "广美";
                    var lastIndex = GetColumnLastRow(maniWs, "存货名称");

                    while (!(currentIndex > lastIndex))
                    {
                        if (moneySum < 113000)
                        {
                            double money;
                            var moneyExist =
                                double.TryParse(
                                    maniWs.Cells[currentIndex, GetColumnByName(maniWs, "含税总金额")].Value?.ToString(),
                                    out money);
                            if (!moneyExist)
                            {
                                money = 0;
                            }

                            moneySum += money;
                            var stockName = maniWs.Cells[currentIndex, GetColumnByName(maniWs, "存货名称")];

                            var nameLength = stockName.Value?.ToString().Length;
                            if (nameLength > 30)
                            {
                                stockName.Value = stockName.Value.ToString().Substring(0, 29);
                            }

                            if (moneySum >= 113000 || currentIndex == lastIndex)
                            {
                                if (moneySum >= 113000)
                                {
                                    moneySum -= money;
                                    currentIndex--;
                                }

                                //生成子表
                                var fi = new FileInfo(outputDir + Path.DirectorySeparatorChar + fileName +
                                                      count.ToString() + ".xlsx");
                                using (ExcelPackage xlPackage = new ExcelPackage(fi))
                                {
                                    //sheet already exist
                                    var xlSheetsList = xlPackage.Workbook.Worksheets;
                                    var destWs = xlSheetsList.Where(x => x.Name == "增值税清单").Count() == 0
                                        ? xlSheetsList.Add("增值税清单", tempWs)
                                        : xlSheetsList["增值税清单"];
                                    string destColumnName;
                                    string srcColumnName;
                                    ExcelRange destRange;
                                    ExcelRange srcRange;
                                    var interval = currentIndex - preIndex + 2;

                                    destColumnName = "货物或应税劳务、服务名称";
                                    srcColumnName = "存货名称";
                                    destRange = GetExcelRange(destWs, 2, interval, destColumnName, destColumnName);
                                    srcRange = GetExcelRange(maniWs, preIndex, currentIndex, srcColumnName,
                                        srcColumnName);
                                    CopyFrom(srcRange, destRange);

                                    destColumnName = "规格型号";
                                    srcColumnName = "存货编码";
                                    destRange = GetExcelRange(destWs, 2, interval, destColumnName, destColumnName);
                                    srcRange = GetExcelRange(maniWs, preIndex, currentIndex, srcColumnName,
                                        srcColumnName);
                                    CopyFrom(srcRange, destRange);

                                    destColumnName = "数量";
                                    srcColumnName = "入库数量";
                                    destRange = GetExcelRange(destWs, 2, interval, destColumnName, destColumnName);
                                    srcRange = GetExcelRange(maniWs, preIndex, currentIndex, srcColumnName,
                                        srcColumnName);
                                    CopyFrom(srcRange, destRange);

                                    destColumnName = "金额";
                                    srcColumnName = "含税总金额";
                                    destRange = GetExcelRange(destWs, 2, interval, destColumnName, destColumnName);
                                    srcRange = GetExcelRange(maniWs, preIndex, currentIndex, srcColumnName,
                                        srcColumnName);
                                    CopyFrom(srcRange, destRange);

                                    destColumnName = "单价";
                                    srcColumnName = "含税单价";
                                    destRange = GetExcelRange(destWs, 2, interval, destColumnName, destColumnName);
                                    srcRange = GetExcelRange(maniWs, preIndex, currentIndex, srcColumnName,
                                        srcColumnName);
                                    CopyFrom(srcRange, destRange);

                                    //copy operation
                                    CopyValue("价格方式", interval, destWs);
                                    //CopyValue("税收分类编码版本号", interval, destWs);
                                    CopyValue("使用优惠政策标识", interval, destWs);
                                    CopyValue("中外合作油气田标识", interval, destWs);

                                    FillValue("税率", interval, destWs, 0.13);
                                    FillValue("税收分类编码", interval, destWs, "1060201070000000000");
                                    FillValue("税收分类编码版本号", interval, destWs, "33.0");
                                    FillValue("计量单位", interval, destWs, "张");

                                    //set index column
                                    var st_lastRow = GetColumnLastRow(destWs, "规格型号");
                                    var st_columnIndex = GetColumnByName(destWs, "序号");
                                    for (int i = 2; i <= st_lastRow; i++)
                                    {
                                        destWs.Cells[i, st_columnIndex].Value = i - 1;
                                    }

                                    preIndex = currentIndex + 1;
                                    xlPackage.SaveAs(fi);
                                    converter.XlsxToXls(fi);
                                }

                                moneySum = 0;
                                count++;
                            }
                        }

                        currentIndex++;
                    }
                }
            }

            if (File.Exists(mani.ToString()))
            {
                File.Delete(mani.ToString());
            }
        }

        /// <summary>
        /// Get column index by column name
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public int GetColumnByName(ExcelWorksheet ws, string columnName)
        {
            if (ws == null) throw new ArgumentNullException(nameof(ws));
            columnName = columnName ?? "0";
            return ws.Cells["1:1"].First(c => c.Value.ToString() == columnName).Start.Column;
        }

        /// <summary>
        /// Get an Excel Range by start row, start column, end row and end column
        /// </summary>
        /// <param name="s"></param>
        /// <param name="rowIndex1"></param>
        /// <param name="rowIndex2"></param>
        /// <param name="columnName1"></param>
        /// <param name="columnName2"></param>
        /// <returns></returns>
        public ExcelRange GetExcelRange(ExcelWorksheet s, int rowIndex1, int rowIndex2, string columnName1,
            string columnName2)
        {
            return s.Cells[rowIndex1, GetColumnByName(s, columnName1), rowIndex2, GetColumnByName(s, columnName2)];
        }

        /// <summary>
        /// Copy a range of cells from source worksheet to destination worksheet
        /// </summary>
        /// <param name="srcWs"></param>
        /// <param name="destWs"></param>
        /// <param name="srcRange"></param>
        /// <param name="destRange"></param>
        public void CopyFrom(ExcelRange srcRange, ExcelRange destRange)
        {
            //bug:out of range
            srcRange.Copy(destRange);
        }

        /// <summary>
        /// Copy value from first row of a column and paste it to fill the column
        /// </summary>
        /// <param name="srcColumnName"></param>
        /// <param name="period">rows of a sub table</param>
        /// <param name="destWs"></param>
        public void CopyValue(string srcColumnName, int interval, ExcelWorksheet destWs)
        {
            var srcColumnIndex = GetColumnByName(destWs, srcColumnName);
            var srcValue = destWs.Cells[2, srcColumnIndex].Value;
            var destRange = destWs.Cells[3, srcColumnIndex, interval, srcColumnIndex];
            destRange.Value = srcValue;
        }

        /// <summary>
        /// Use given value to fill the column
        /// </summary>
        /// <param name="srcColumnName"></param>
        /// <param name="interval">rows of a sub table</param>
        /// <param name="destWs"></param>
        /// <param name="srcValue"> Value given to be filled</param>
        public void FillValue(string srcColumnName, int interval, ExcelWorksheet destWs, object srcValue)
        {
            var srcColumnIndex = GetColumnByName(destWs, srcColumnName);
            var destRange = destWs.Cells[2, srcColumnIndex, interval, srcColumnIndex];
            destRange.Value = srcValue;
        }

        /// <summary>
        /// Get last row index of a column
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public int GetColumnLastRow(ExcelWorksheet ws, string columnName)
        {
            var columnIndex = GetColumnByName(ws, columnName);
            var lastIndex = ws.Cells[ws.Dimension.Start.Row, columnIndex, ws.Dimension.End.Row, columnIndex]
                .Last(c => c.Value.ToString() != "").End.Row;
            return lastIndex;
        }
    }
}