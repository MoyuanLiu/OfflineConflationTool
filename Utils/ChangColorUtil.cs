using NPOI.XSSF.UserModel;
using System.Drawing;

namespace Utils
{
    public class ChangColorUtil
    {
        public static XSSFSheet ChangeColor(XSSFSheet sheet, int x, int y, Color color,XSSFCellStyle cellstyle)
        {
            XSSFRow row = (XSSFRow)sheet.GetRow(x);
            XSSFCell cell = (XSSFCell)row.GetCell(y);
            XSSFColor XlColour = new XSSFColor(color);
            cellstyle.SetFillForegroundColor(XlColour);
            cellstyle.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
            cell.CellStyle = cellstyle;
            return sheet;
        }
        
    }
}
