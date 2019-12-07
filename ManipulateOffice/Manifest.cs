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

//Todo:Negative numbers
namespace Manifest
{
    class ReportGenerate
    {
        public string ManifestFileName { get; set; }
        public string TemplateFileName { get; set; } = "template.xlsx";
        public string ManifestFileDir { get; set; } 
        public string TemplateFileDir { get; set; } = @"D:\清单模板";

        public string TotalAmountWithTax { get; set; } = "含税总金额";

        public string InventoryName { get; set; } = "存货名称";

        public string InventoryCode { get; set; } = "存货编码";

        public string InventoryAmount { get; set; } = "入库数量";

        public string UnitWithTax { get; set; } = "含税单价";
        public ExcelTypeConverter converter { get; set; }

        public ReportGenerate()
        {
            converter = new ExcelTypeConverter();
        }

        public void ExcelGenearte()
        {
            var temp = new FileInfo(TemplateFileDir + Path.DirectorySeparatorChar + TemplateFileName);
            var originManiName = ManifestFileDir + Path.DirectorySeparatorChar + ManifestFileName;
            FileInfo mani;

            if (ManifestFileName.EndsWith(".xls"))
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
                if (!ColumnNameExist(maniWs, InventoryName))
                {
                    throw new InvalidOperationException($"列{InventoryName}未配置");
                }
                if (!ColumnNameExist(maniWs, TotalAmountWithTax))
                {
                    throw new InvalidOperationException($"列{TotalAmountWithTax}未配置");
                }
                if (!ColumnNameExist(maniWs, InventoryAmount))
                {
                    throw new InvalidOperationException($"列{InventoryAmount}未配置");
                }
                if (!ColumnNameExist(maniWs, UnitWithTax))
                {
                    throw new InvalidOperationException($"列{UnitWithTax}未配置");
                }

                //To Do clean up the Data
                using (ExcelPackage tempPackage = new ExcelPackage(temp))
                {
                    //增值税清单模板
                    var tempWs = tempPackage.Workbook.Worksheets.FirstOrDefault();
                    int currentIndex = 2;
                    int count = 1;
                    int preIndex = 2;
                    double moneySum = 0;
                    string outputDir = @"D:\清单结果";
                    string fileName = ManifestFileName.Split('.')[0] + "子表";
                    var lastIndex = GetColumnLastRow(maniWs, InventoryName);
                    while (!(currentIndex > lastIndex))
                    {
                        if (moneySum < 113000)
                        {
                            double money;
                            var moneyExist =
                                double.TryParse(
                                    maniWs.Cells[currentIndex, GetColumnByName(maniWs, TotalAmountWithTax)].Value?.ToString(),
                                    out money);
                            if (!moneyExist)
                            {
                                money = 0;
                            }
                            else if(money < 0)
                            {
                                throw new InvalidOperationException($"列{TotalAmountWithTax}不能为负数");
                            }

                            moneySum += money;
                            var stockName = maniWs.Cells[currentIndex, GetColumnByName(maniWs, InventoryName)];

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
                                    var destWs = xlSheetsList.All(x => x.Name != "增值税清单")
                                        ? xlSheetsList.Add("增值税清单", tempWs)
                                        : xlSheetsList["增值税清单"];
                                    var interval = currentIndex - preIndex + 2;

                                    string destColumnName = "货物或应税劳务、服务名称";
                                    string srcColumnName = InventoryName;
                                    ExcelRange destRange = GetExcelRange(destWs, 2, interval, destColumnName, destColumnName);
                                    ExcelRange srcRange = GetExcelRange(maniWs, preIndex, currentIndex, srcColumnName,
                                        srcColumnName);
                                    CopyFrom(srcRange, destRange);

                                    if (ColumnNameExist(maniWs, InventoryCode))
                                    {
                                        destColumnName = "规格型号";
                                        srcColumnName = InventoryCode;
                                        destRange = GetExcelRange(destWs, 2, interval, destColumnName, destColumnName);
                                        srcRange = GetExcelRange(maniWs, preIndex, currentIndex, srcColumnName,
                                            srcColumnName);
                                        CopyFrom(srcRange, destRange);
                                    }

                                    destColumnName = "数量";
                                    srcColumnName = InventoryAmount;
                                    destRange = GetExcelRange(destWs, 2, interval, destColumnName, destColumnName);
                                    srcRange = GetExcelRange(maniWs, preIndex, currentIndex, srcColumnName,
                                        srcColumnName);
                                    CopyFrom(srcRange, destRange);

                                    destColumnName = "金额";
                                    srcColumnName = TotalAmountWithTax;
                                    destRange = GetExcelRange(destWs, 2, interval, destColumnName, destColumnName);
                                    srcRange = GetExcelRange(maniWs, preIndex, currentIndex, srcColumnName,
                                        srcColumnName);
                                    CopyFrom(srcRange, destRange);

                                    destColumnName = "单价";
                                    srcColumnName = UnitWithTax;
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
                                    var st_lastRow = GetColumnLastRow(destWs, "金额");
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
        /// Get column index by column name，return bool
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        private bool ColumnNameExist(ExcelWorksheet ws, string columnName)
        {
            if (ws == null) throw new ArgumentNullException(nameof(ws));
            return ws.Cells["1:1"].FirstOrDefault(c => c.Value.ToString() == columnName) != null;
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
            columnName = columnName ?? "";
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
        private ExcelRange GetExcelRange(ExcelWorksheet s, int rowIndex1, int rowIndex2, string columnName1,
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
        private void CopyFrom(ExcelRange srcRange, ExcelRange destRange)
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
        private void CopyValue(string srcColumnName, int interval, ExcelWorksheet destWs)
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
        private void FillValue(string srcColumnName, int interval, ExcelWorksheet destWs, object srcValue)
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
        private int GetColumnLastRow(ExcelWorksheet ws, string columnName)
        {
            var columnIndex = GetColumnByName(ws, columnName);
            var lastIndex = ws.Cells[ws.Dimension.Start.Row, columnIndex, ws.Dimension.End.Row, columnIndex]
                .Last(c => c.Value.ToString() != "").End.Row;
            return lastIndex;
        }
    }
}