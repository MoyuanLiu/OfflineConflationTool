using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace Utils
{
    public class ExcelDataTableUtil
    {
        public static XSSFWorkbook LoadExcel(string filename)
        {
            XSSFWorkbook xssfwb;
            using (FileStream file = new FileStream(filename, FileMode.Open, FileAccess.Read))
            {
                xssfwb = new XSSFWorkbook(file);
            }
            return xssfwb;
        }
        /// <summary>
        /// Convert Excel sheets to DataTable list
        /// </summary>
        /// <param name="filename"></param>
        /// <returns>list of datatable</returns>
        public static List<DataTable> ExceltoDataTable(string filename)
        {
            XSSFWorkbook xssfwb;
            List<DataTable> dts = new List<DataTable>();
            using (FileStream file = new FileStream(filename, FileMode.Open, FileAccess.Read))
            {
                xssfwb = new XSSFWorkbook(file);
            }
            for (int i = 0; i < xssfwb.NumberOfSheets; i++)
            {
                XSSFSheet sheet = (XSSFSheet)xssfwb.GetSheetAt(i);
                DataTable dt = new DataTable();
                int num = 0;
                while (sheet.GetRow(num) != null)
                {
                    if (dt.Columns.Count < sheet.GetRow(num).Cells.Count)
                    {
                        for (int j = 0; j < sheet.GetRow(num).Cells.Count; j++)
                        {
                            dt.Columns.Add("", typeof(string));
                        }
                    }
                    XSSFRow row = (XSSFRow)sheet.GetRow(num);
                    
                    DataRow dr = dt.Rows.Add();

                    for (int k = 0; k < row.Cells.Count; k++)
                    {

                        XSSFCell cell = (XSSFCell)row.GetCell(k);

                        if (cell != null)
                        {
                            switch (cell.CellType)
                            {
                                case CellType.Numeric:
                                    dr[k] = cell.NumericCellValue;
                                    break;
                                case CellType.String:
                                    dr[k] = cell.StringCellValue;
                                    break;
                                case CellType.Blank:
                                    dr[k] = "";
                                    break;
                                case CellType.Boolean:
                                    dr[k] = cell.BooleanCellValue;
                                    break;
                            }
                        }


                    }
                    num++;
                }
                dts.Add(dt);
            }
            return dts;
        }
        public static DataTable SheetToDataTable(XSSFSheet sheet,int firstrow,int lastrow)
        {
            DataTable dt = new DataTable();
            
            for (int i = firstrow;i<=lastrow;i++)
            {
                if (dt.Columns.Count < sheet.GetRow(i).Cells.Count)
                {
                    for (int j = 0; j < sheet.GetRow(i).Cells.Count; j++)
                    {
                        dt.Columns.Add("", typeof(string));
                    }
                }
            }
            for (int j = firstrow; j <= lastrow; j++)
            {
                XSSFRow row = (XSSFRow)sheet.GetRow(j);
                DataRow dr = dt.Rows.Add();
                for (int k = 0; k < row.Cells.Count; k++)
                {

                    XSSFCell cell = (XSSFCell)row.GetCell(k);
                    
                    if (cell != null)
                    {
                        switch (cell.CellType)
                        {
                            case CellType.Numeric:
                                dr[k] = cell.NumericCellValue;
                                break;
                            case CellType.String:
                                dr[k] = cell.StringCellValue;
                                break;
                            case CellType.Blank:
                                dr[k] = "";
                                break;
                            case CellType.Boolean:
                                dr[k] = cell.BooleanCellValue;
                                break;
                        }
                    }
                }
            }
            return dt;
        }

        public static List<XSSFSheet> GetAllSheets(string filename)
        {
            List<XSSFSheet> sheets = new List<XSSFSheet>();
            XSSFWorkbook xssfwb;
            using (FileStream file = new FileStream(filename, FileMode.Open, FileAccess.Read))
            {
                xssfwb = new XSSFWorkbook(file);
            }
            for (int i = 0; i < xssfwb.NumberOfSheets; i++)
            {
                XSSFSheet sheet = (XSSFSheet)xssfwb.GetSheetAt(i);
                sheets.Add(sheet);
            }
            return sheets;
        }
        public static XSSFSheet GetSheetbyName(string filename, string sheetname)
        {
            int index = 0;
            XSSFWorkbook xssfwb;
            using (FileStream file = new FileStream(filename, FileMode.Open, FileAccess.Read))
            {
                xssfwb = new XSSFWorkbook(file);
            }
            for (int i = 0; i < xssfwb.NumberOfSheets; i++)
            {
                XSSFSheet sheet = (XSSFSheet)xssfwb.GetSheetAt(i);
                if (sheet.SheetName== sheetname)
                {
                    index = i;
                }
            }
            return (XSSFSheet)xssfwb.GetSheetAt(index);
        }
        public static XSSFWorkbook DataTabletoExcel(DataTable dt)
        {
            XSSFWorkbook xssfwb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)xssfwb.CreateSheet();
            sheet.CreateRow(dt.Rows.Count);
            for (int i = 0;i< dt.Rows.Count;i++)
            {
                DataRow dr = dt.Rows[i];
                sheet.CreateRow(i);
                XSSFRow row = (XSSFRow)sheet.GetRow(i);
                for (int j =0;j< dt.Columns.Count;j++)
                {
                    XSSFCell cell = (XSSFCell)row.CreateCell(j);
                    cell.SetCellValue(dr[j].ToString());
                }
            }
            return xssfwb;
        }
        public static void WriteExcel(XSSFWorkbook xssfwb, string target)
        {
            FileStream sw = File.Create(target);

            xssfwb.Write(sw);

            sw.Close();
        } 
    
}
}
