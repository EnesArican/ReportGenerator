using ReportGenerator.Interfaces;
using Microsoft.Office.Interop.Excel;
using C = ReportGenerator.Constants.ColourScheme;
using SC = ReportGenerator.Constants.SystemConstant;

namespace ReportGenerator.Services
{
    public class FormatterService : IFormatterService
    {
        private readonly IRangeEditorService _rangeEditor;

        public FormatterService(IRangeEditorService rangeEditor) 
        {
            _rangeEditor = rangeEditor;
        }
        public void FormatWorksheet(Worksheet sheet, int AtColsCount)
        {
            var atEndCol = SC.AtStrtCol + AtColsCount - 1;
            var atEndRow = sheet.UsedRange.Rows.Count;

            AddColumHeaders(sheet, atEndCol);
            _rangeEditor.AddBorders(sheet.UsedRange);

            AddSpecialCellFormats(sheet);

            sheet.UsedRange.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            sheet.UsedRange.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;

            MergeMonthHeader(sheet, atEndCol);

            FormatAttendanceRange(sheet, atEndRow, atEndCol);

            FormatNameAndMobileColsRange(sheet, atEndRow);

            FormatNumbersRange(sheet, atEndRow);

            FormatAllHeaderRows(sheet);

            FormatTotalColsRange(sheet);
        }

        private void AddColumHeaders(Worksheet sheet, int atEndCol) 
        {
            sheet.Cells[1, 1] = "Yoklama Ayi";
            sheet.Cells[SC.DateRow, 1] = "Yoklama Tarihleri";

            sheet.Cells[1, atEndCol + 1] = "Ay sonu iştirak durumu";
            sheet.Cells[SC.DateRow, atEndCol + 1] = "İştirak etmesi gereken";
            sheet.Cells[SC.DateRow, atEndCol + 2] = "İştirak ettiği";
            sheet.Cells[SC.DateRow, atEndCol + 3] = "Nisbeti";

            sheet.Cells[SC.HeaderRow, 1] = "SN.";
            sheet.Cells[SC.HeaderRow, 2] = "Adı";
            sheet.Cells[SC.HeaderRow, 3] = "Soyadı";
            sheet.Cells[SC.HeaderRow, 4] = "Telefonu";
        }

        private void AddSpecialCellFormats(Worksheet sheet) 
        {
            Range mobileColumn = (Range)sheet.Columns[SC.MobileCol];
            mobileColumn.NumberFormat = "#### ### #### ###";

            Range dateRowRange = (Range)sheet.Rows[SC.DateRow];
            dateRowRange.NumberFormat = "dd-MM-yyyy";

            Range AtPercentCol = (Range)sheet.Columns[sheet.UsedRange.Columns.Count];
            AtPercentCol.NumberFormat = "0%";
        }

        private void MergeMonthHeader(Worksheet sheet, int atEndCol) 
        {
            var strtCell = sheet.Cells[1, SC.AtStrtCol];
            var endCell = sheet.Cells[1, atEndCol];
            var monthHeader = sheet.Range[strtCell, endCell];
            monthHeader.Merge();

            strtCell = sheet.Cells[1, 1];
            endCell = sheet.Cells[1, SC.MobileCol];
            var monthTitle = sheet.Range[strtCell, endCell];
            monthTitle.Merge();

            strtCell = sheet.Cells[SC.DateRow, 1];
            endCell = sheet.Cells[SC.DateRow, SC.MobileCol];
            var datesTitle = sheet.Range[strtCell, endCell];
            datesTitle.Merge();



        }

        private void FormatAttendanceRange(Worksheet sheet, int atEndRow, int atEndCol) 
        {
            var attendanceRange = _rangeEditor.GetRange(sheet, SC.DataStrtRow, SC.AtStrtCol, atEndRow, atEndCol);
            attendanceRange.Borders[XlBordersIndex.xlInsideVertical].Weight = XlBorderWeight.xlMedium;
            _rangeEditor.AddFormatConditions(attendanceRange);
        }

        private void FormatNameAndMobileColsRange(Worksheet sheet, int atEndRow) 
        {
            var nameAndMobileColsRange = _rangeEditor.GetRange(sheet, SC.DataStrtRow, SC.FirstNameCol, atEndRow, SC.MobileCol);
            nameAndMobileColsRange.Columns.AutoFit();
            nameAndMobileColsRange.Cells.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            _rangeEditor.AddAlternateRowColours(nameAndMobileColsRange);
        }

        private void FormatNumbersRange(Worksheet sheet, int atEndRow) 
        {
            var numbersRange = _rangeEditor.GetRange(sheet, SC.DataStrtRow, 1, atEndRow, 1);
            numbersRange.ColumnWidth = 3.5;
            _rangeEditor.FormatRange(numbersRange, C.Cream, C.Red, 10, true);
        }

        private void FormatAllHeaderRows(Worksheet sheet)
        {
            var titleRow = (Range)sheet.UsedRange.Rows[1];
            _rangeEditor.FormatRange(titleRow, C.Cream, C.Red, 13, true);

            var dateRow = (Range)sheet.UsedRange.Rows[SC.DateRow];
            _rangeEditor.FormatRange(dateRow, C.Cream, C.Navy, 9, true);

            var headerRow = (Range)sheet.UsedRange.Rows[SC.HeaderRow];
            _rangeEditor.FormatRange(headerRow, C.Cream, C.Navy, 9, true);
        }

        private void FormatTotalColsRange(Worksheet sheet) 
        { 
        
        }

    }
}
