using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace StudentInfoSystem.Models
{
    public enum TextAlign
    {
        Center = 0,
        Left = 1,
        Right = 2
    }

    public class TableHeader
    {
        public TableHeader()
        {
            Header = string.Empty;
            Width = 0f;
            IsRotate = false;
            ChildHeader = new List<TableHeader>();
            Align = TextAlign.Center;
            IsBold = false;
        }
        public string Header { get; set; }
        public string Weight { get; set; }
        public float Width { get; set; }
        public bool IsRotate { get; set; }
        public TextAlign Align { get; set; }
        public List<TableHeader> ChildHeader { get; set; }
        public bool IsBold { get; set; }
    }

    public class ExcelTool
    {
        public static string Formatter = "";
        private static void CellFill(TableHeader oTableHeader, bool IsChild, ref ExcelRange cell, ref ExcelWorksheet sheet, ref int nRowIndex, int nStartCol, int nEndCol, int fontSize)
        {
            OfficeOpenXml.Style.Border border;

            if (oTableHeader.IsRotate)
                cell.Style.TextRotation = 90;

            cell.Value = oTableHeader.Header; cell.Style.Font.Bold = false; cell.Style.WrapText = true; cell.Style.Font.Size = fontSize;

            if (TextAlign.Center == oTableHeader.Align)
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            else if (TextAlign.Left == oTableHeader.Align)
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            else if (TextAlign.Right == oTableHeader.Align)
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            else
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            border = cell.Style.Border; border.Bottom.Style = border.Top.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
        }
        public static ExcelRange FillCellBasic(ExcelWorksheet sheet, int nRowIndex, int nStartCol, string sVal, bool IsNumber, ExcelHorizontalAlignment oExcelHorizontalAlignment, bool IsBold, bool isGray)
        {
            ExcelRange cell;
            OfficeOpenXml.Style.Border border;

            cell = sheet.Cells[nRowIndex, nStartCol++];
            if (IsNumber)
            {
                cell.Value = Convert.ToDouble(sVal);
            }
            else
            {
                cell.Value = sVal;
            }
            cell.Style.Font.Bold = IsBold;
            cell.Style.WrapText = true;
            if (IsNumber)
            {
                cell.Style.Numberformat.Format = Formatter;
            }
            if (isGray)
            {
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid; cell.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            }
            cell.Style.HorizontalAlignment = oExcelHorizontalAlignment;
            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            border = cell.Style.Border; border.Bottom.Style = border.Top.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
            return cell;
        }
        public static ExcelRange FillCellBasic(ExcelWorksheet sheet, int nRowIndex, int nStartCol, string sVal, bool IsNumber, ExcelHorizontalAlignment oExcelHorizontalAlignment, bool IsBold, Color oColor)
        {
            ExcelRange cell;
            OfficeOpenXml.Style.Border border;

            cell = sheet.Cells[nRowIndex, nStartCol++];
            if (IsNumber)
            {
                cell.Value = Convert.ToDouble(sVal);
            }
            else
            {
                cell.Value = sVal;
            }
            cell.Style.Font.Bold = IsBold;
            cell.Style.WrapText = true;
            if (IsNumber)
            {
                cell.Style.Numberformat.Format = Formatter;
            }

            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(oColor);

            cell.Style.HorizontalAlignment = oExcelHorizontalAlignment;
            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            border = cell.Style.Border; border.Bottom.Style = border.Top.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
            return cell;
        }
        private static void CellFill(TableHeader oTableHeader, bool IsChild, ref ExcelRange cell, ref ExcelWorksheet sheet, ref int nRowIndex, int nStartCol, int nEndCol, int fontSize, bool isBold, bool isGray)
        {
            OfficeOpenXml.Style.Border border;

            if (oTableHeader.IsRotate)
                cell.Style.TextRotation = 90;

            cell.Value = oTableHeader.Header; cell.Style.Font.Bold = isBold; cell.Style.WrapText = true; cell.Style.Font.Size = fontSize;


            if (TextAlign.Center == oTableHeader.Align)
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            else if (TextAlign.Left == oTableHeader.Align)
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            else if (TextAlign.Right == oTableHeader.Align)
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            else
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            if (isGray)
            {
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid; cell.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            }

            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            border = cell.Style.Border; border.Bottom.Style = border.Top.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
        }

        public static void GenerateHeader(List<TableHeader> table_header, ref ExcelWorksheet sheet, ref int nRowIndex, int nStartCol, int nEndCol, int fontSize)
        {
            ExcelRange cell;
            #region Report Header
            fontSize = (fontSize <= 0) ? 10 : fontSize; // default font size
            bool hasChild = table_header.Exists(x => x.ChildHeader.Count > 0);
            int span = 0;
            foreach (TableHeader listItem in table_header)
            {
                span = 0;
                if (hasChild)
                {
                    span = listItem.ChildHeader.Count;
                    if (span > 0)
                    {
                        cell = sheet.Cells[nRowIndex, nStartCol, nRowIndex, nStartCol + span - 1];
                        cell.Merge = true;
                        nStartCol = nStartCol + span;
                    }
                    else
                    {
                        cell = sheet.Cells[nRowIndex, nStartCol, nRowIndex + 1, nStartCol++];
                        cell.Merge = true;
                    }
                }
                else
                {
                    cell = sheet.Cells[nRowIndex, nStartCol++];

                }

                CellFill(listItem, false, ref cell, ref sheet, ref nRowIndex, nStartCol, nEndCol, fontSize);

                // If any child header found
                if (span > 0)
                {
                    nStartCol -= span;
                    foreach (TableHeader childHeader in listItem.ChildHeader)
                    {
                        cell = sheet.Cells[nRowIndex + 1, nStartCol, nRowIndex + 1, nStartCol++];
                        CellFill(childHeader, true, ref cell, ref sheet, ref nRowIndex, nStartCol, nEndCol, fontSize);
                    }
                }
            }

            nRowIndex += (hasChild) ? 2 : 1;
            #endregion
        }
        public static void GenerateHeader(List<TableHeader> table_header, ref ExcelWorksheet sheet, ref int nRowIndex, int nStartCol, int nEndCol, int fontSize, bool isBold, bool isGray)
        {
            ExcelRange cell;
            #region Report Header
            fontSize = (fontSize <= 0) ? 10 : fontSize; // default font size
            bool hasChild = table_header.Exists(x => x.ChildHeader.Count > 0);
            int span = 0;
            foreach (TableHeader listItem in table_header)
            {
                span = 0;
                if (hasChild)
                {
                    span = listItem.ChildHeader.Count;
                    if (span > 0)
                    {
                        cell = sheet.Cells[nRowIndex, nStartCol, nRowIndex, nStartCol + span - 1];
                        cell.Merge = true;
                        nStartCol = nStartCol + span;
                    }
                    else
                    {
                        cell = sheet.Cells[nRowIndex, nStartCol, nRowIndex + 1, nStartCol++];
                        cell.Merge = true;
                    }
                }
                else
                {
                    cell = sheet.Cells[nRowIndex, nStartCol++];

                }

                CellFill(listItem, false, ref cell, ref sheet, ref nRowIndex, nStartCol, nEndCol, fontSize, isBold, isGray);

                // If any child header found
                if (span > 0)
                {
                    nStartCol -= span;
                    foreach (TableHeader childHeader in listItem.ChildHeader)
                    {
                        cell = sheet.Cells[nRowIndex + 1, nStartCol, nRowIndex + 1, nStartCol++];
                        CellFill(childHeader, true, ref cell, ref sheet, ref nRowIndex, nStartCol, nEndCol, fontSize, isBold, isGray);
                    }
                }
            }

            nRowIndex += (hasChild) ? 2 : 1;
            #endregion
        }
        public static void SetColumnWidth(List<TableHeader> table_header, ref ExcelWorksheet sheet, ref int nStartCol, ref int nEndCol)
        {
            foreach (TableHeader listItem in table_header)
            {
                if (listItem.ChildHeader.Count > 0)
                {
                    foreach (TableHeader child in listItem.ChildHeader)
                    {
                        nEndCol++;
                        sheet.Column(nStartCol++).Width = child.Width;
                    }
                }
                else
                {
                    nEndCol++;
                    sheet.Column(nStartCol++).Width = listItem.Width;
                }
            }
        }

        public static ExcelRange FillCell(ExcelWorksheet sheet, int nRowIndex, int nStartCol, string sVal, bool IsNumber)
        {
            return FillCell(sheet, nRowIndex, nStartCol, sVal, IsNumber, false);
        }

        public static ExcelRange FillCell(ExcelWorksheet sheet, int nRowIndex, int nStartCol, string sVal, bool IsNumber, bool IsBold)
        {
            ExcelRange cell;
            OfficeOpenXml.Style.Border border;

            cell = sheet.Cells[nRowIndex, nStartCol++];
            if (IsNumber)
                cell.Value = Convert.ToDouble(sVal);
            else
                cell.Value = sVal;
            cell.Style.Font.Bold = IsBold;
            cell.Style.WrapText = true;
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            if (IsNumber)
            {
                cell.Style.Numberformat.Format = Formatter;
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            }
            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            border = cell.Style.Border; border.Bottom.Style = border.Top.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
            return cell;
        }
        public static ExcelRange FillCell(ExcelWorksheet sheet, int nRowIndex, int nStartCol, string sVal, bool IsNumber, bool IsBold, ExcelHorizontalAlignment ePosition, bool IsGray)
        {
            ExcelRange cell;
            OfficeOpenXml.Style.Border border;

            cell = sheet.Cells[nRowIndex, nStartCol++];
            if (IsNumber)
            {
                cell.Value = Convert.ToDouble(sVal);
                cell.Style.Numberformat.Format = Formatter;
            }
            else
            {
                cell.Value = sVal;
                cell.Style.WrapText = true;
            }
            if (IsGray) { cell.Style.Fill.PatternType = ExcelFillStyle.Solid; cell.Style.Fill.BackgroundColor.SetColor(Color.LightGray); }
            if (IsBold) { cell.Style.Font.Bold = IsBold; }
            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            cell.Style.HorizontalAlignment = ePosition;
            border = cell.Style.Border; border.Bottom.Style = border.Top.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
            return cell;
        }
        public static ExcelRange FillCellOne(ExcelWorksheet sheet, int nRowIndex, int nStartCol, string sVal, int bPosition, bool IsBold, int fontSize, bool isGray, bool LeftBorder, bool RightBorder, bool UpBorder, bool DownBorder)
        {
            ExcelRange cell;
            OfficeOpenXml.Style.Border border;

            cell = sheet.Cells[nRowIndex, nStartCol++];
            cell.Value = sVal;
            cell.Style.Font.Size = fontSize;
            cell.Style.Font.Bold = IsBold;
            cell.Style.WrapText = true;
            if (bPosition == 1)
            {
                cell.Style.Numberformat.Format = Formatter;
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            }
            if (bPosition == 2)
            {
                cell.Style.Numberformat.Format = Formatter;
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            if (bPosition == 3)
            {
                cell.Style.Numberformat.Format = Formatter;
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            }
            if (isGray)
            {
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid; cell.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            }
            border = cell.Style.Border; border.Bottom.Style = border.Top.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
            if (!LeftBorder) border.Left.Style = ExcelBorderStyle.None;
            if (!RightBorder) border.Right.Style = ExcelBorderStyle.None;
            if (!UpBorder) border.Top.Style = ExcelBorderStyle.None;
            if (!DownBorder) border.Bottom.Style = ExcelBorderStyle.None;
            return cell;
        }
        public static void FillCellMerge(ref ExcelWorksheet sheet, int startRowIndex, int endRowIndex, int startColIndex, int endColIndex, string sVal)
        {
            ExcelRange cell;
            OfficeOpenXml.Style.Border border;

            cell = sheet.Cells[startRowIndex, startColIndex, endRowIndex, endColIndex];
            cell.Merge = true;
            cell.Value = sVal;
            cell.Style.Font.Bold = false;
            cell.Style.WrapText = true;
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            border = cell.Style.Border; border.Bottom.Style = border.Top.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
        }
        public static void FillCellMerge(ref ExcelWorksheet sheet, string sVal, int startRowIndex, int endRowIndex, int startColIndex, int endColIndex)
        {
            FillCellMerge(ref sheet, sVal, startRowIndex, endRowIndex, startColIndex, endColIndex, false, ExcelHorizontalAlignment.Left);
        }
        public static void FillCellMerge(ref ExcelWorksheet sheet, string sVal, int startRowIndex, int endRowIndex, int startColIndex, int endColIndex, bool isBold, ExcelHorizontalAlignment HoriAlign)
        {
            FillCellMerge(ref sheet, sVal, startRowIndex, endRowIndex, startColIndex, endColIndex, isBold, HoriAlign, false);
        }
        public static void FillCellMerge(ref ExcelWorksheet sheet, string sVal, int startRowIndex, int endRowIndex, int startColIndex, int endColIndex, bool isBold, ExcelHorizontalAlignment HoriAlign, bool isGray)
        {
            ExcelRange cell;
            OfficeOpenXml.Style.Border border;

            cell = sheet.Cells[startRowIndex, startColIndex, endRowIndex, endColIndex];
            cell.Merge = true;
            cell.Value = sVal;
            cell.Style.Font.Bold = isBold;
            cell.Style.WrapText = true;
            cell.Style.HorizontalAlignment = HoriAlign;
            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            if (isGray)
            {
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid; cell.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            }
            border = cell.Style.Border; border.Bottom.Style = border.Top.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
        }
        public static void FillCellMerge(ref ExcelWorksheet sheet, string sVal, int startRowIndex, int endRowIndex, int startColIndex, int endColIndex, bool isBold, ExcelHorizontalAlignment HoriAlign, Color oColor)
        {
            ExcelRange cell;
            OfficeOpenXml.Style.Border border;

            cell = sheet.Cells[startRowIndex, startColIndex, endRowIndex, endColIndex];
            cell.Merge = true;
            cell.Value = sVal;
            cell.Style.Font.Bold = isBold;
            cell.Style.WrapText = true;
            cell.Style.HorizontalAlignment = HoriAlign;
            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid; cell.Style.Fill.BackgroundColor.SetColor(oColor);

            border = cell.Style.Border; border.Bottom.Style = border.Top.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
        }
        public static void FillCellMerge(ref ExcelWorksheet sheet, double dVal, int startRowIndex, int endRowIndex, int startColIndex, int endColIndex, bool isBold)
        {
            ExcelRange cell;
            OfficeOpenXml.Style.Border border;

            cell = sheet.Cells[startRowIndex, startColIndex, endRowIndex, endColIndex];
            cell.Merge = true;
            cell.Value = dVal;
            cell.Style.Font.Bold = isBold;
            cell.Style.WrapText = true;
            cell.Style.Numberformat.Format = Formatter;
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            border = cell.Style.Border; border.Bottom.Style = border.Top.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
        }

        public static void FillCellMerge(ref ExcelWorksheet sheet, string sVal, int startRowIndex, int endRowIndex, int startColIndex, int endColIndex, bool isBold, ExcelHorizontalAlignment excelHorizontalAlignment, ExcelVerticalAlignment excelVerticalAlignment, bool isGray)
        {
            ExcelRange cell;
            OfficeOpenXml.Style.Border border;

            cell = sheet.Cells[startRowIndex, startColIndex, endRowIndex, endColIndex];
            cell.Merge = true;
            cell.Value = sVal;
            cell.Style.Font.Bold = isBold;
            cell.Style.WrapText = true;
            cell.Style.HorizontalAlignment = excelHorizontalAlignment;
            cell.Style.VerticalAlignment = excelVerticalAlignment;
            if (isGray)
            {
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid; cell.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            }
            border = cell.Style.Border; border.Bottom.Style = border.Top.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
        }

        public static void FillCellMergeForNumber(ref ExcelWorksheet sheet, double sVal, int startRowIndex, int endRowIndex, int startColIndex, int endColIndex, bool isBold, ExcelHorizontalAlignment excelHorizontalAlignment, ExcelVerticalAlignment excelVerticalAlignment, bool isGray)
        {
            ExcelRange cell;
            OfficeOpenXml.Style.Border border;

            cell = sheet.Cells[startRowIndex, startColIndex, endRowIndex, endColIndex];
            cell.Merge = true;
            cell.Value = sVal;
            cell.Style.Font.Bold = isBold;
            cell.Style.WrapText = true;
            cell.Style.HorizontalAlignment = excelHorizontalAlignment;
            cell.Style.VerticalAlignment = excelVerticalAlignment;
            cell.Style.Numberformat.Format = Formatter;
            if (isGray)
            {
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid; cell.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            }
            border = cell.Style.Border; border.Bottom.Style = border.Top.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
        }

        public static void FillCellMergeBasic(ref ExcelWorksheet sheet, string sVal, bool isNumber, int startRowIndex, int endRowIndex, int startColIndex, int endColIndex, bool isBold, ExcelHorizontalAlignment HoriAlign, bool isGray)
        {
            ExcelRange cell;
            OfficeOpenXml.Style.Border border;

            cell = sheet.Cells[startRowIndex, startColIndex, endRowIndex, endColIndex];
            cell.Merge = true;
            if (isNumber)
            {
                cell.Value = Convert.ToDouble(sVal);
            }
            else
            {
                cell.Value = sVal;
            }
            cell.Style.Font.Bold = isBold;
            cell.Style.WrapText = true;
            cell.Style.HorizontalAlignment = HoriAlign;
            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            //if (isGray)
            //{
            //    cell.Style.Fill.PatternType = ExcelFillStyle.Solid; cell.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            //}
            border = cell.Style.Border; border.Bottom.Style = border.Top.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
        }

    }
}
