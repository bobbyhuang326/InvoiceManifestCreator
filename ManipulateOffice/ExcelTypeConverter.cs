using System.IO;
using Spire.Xls;

namespace Manifest
{
    //todo: due to trail problem, rewrite this by NPOI
    class ExcelTypeConverter
    {
        /// <summary>
        /// Convert xls to xlsx file
        /// </summary>
        /// <param name="filesFolder"></param>
        public bool XlsToXlsx(FileInfo file)
        {
            var fileName = file.ToString();
            if (!File.Exists(fileName))
            {
                return false;
            }
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(fileName);
            var convertedName = fileName + "x";
            workbook.SaveToFile(convertedName, ExcelVersion.Version2013);
            return true;
        }

        public bool XlsxToXls(FileInfo file)
        {
            var fileName = file.ToString();
            if (!File.Exists(fileName))
            {
                return false;
            }
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(fileName);
            var convertedName = fileName.Remove(fileName.Length-1);
            workbook.SaveToFile(convertedName, ExcelVersion.Version97to2003);
            File.Delete(fileName);
            return true;
        }
    }
}