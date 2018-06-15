// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
using System.Reflection;

namespace Excel.IO
{
    public static class ExcelExtensions
    {
        // These numbers are the format ids that correspond to OA dates in Excel/OOXML Spreadsheets
        private static int[] dateNumberFormats = new int[] { 14, 15, 16, 17, 22 };

        /// <summary>
        /// Finds the column identifier for a given cell, ie: A
        /// </summary>
        /// <param name="cell">The Cell to find the column for</param>
        /// <returns>The column name</returns>
        public static string GetColumn(this Cell cell)
        {
            if (!cell.CellReference.HasValue)
            {
                return string.Empty;
            }

            var endIndex = 0;

            for (int i = 0; i < cell.CellReference.Value.Length; i++)
            {
                if (char.IsLetter(cell.CellReference.Value[i]))
                {
                    endIndex = i + 1;
                }
            }

            return cell.CellReference.Value.Substring(0, endIndex);
        }

        /// <summary>
        /// Returns the value for a given Cell, taking into account the shared string table
        /// </summary>
        /// <param name="cell">The Cell to get the value for</param>
        /// <returns>The value of the Cell or null</returns>
        public static string GetCellValue(this Cell cell)
        {
            if (cell == null)
            {
                return null;
            }

            if (string.IsNullOrWhiteSpace(cell.DataType))
            {
                var dateString = string.Empty;

                if (cell.TryParseDate(out dateString))
                {
                    return dateString;
                }

                return cell.InnerText;
            }

            switch (cell.DataType.Value)
            {
                case CellValues.SharedString:
                    {
                        var worksheet = cell.FindParentWorksheet();
                        var sharedStringTablePart = worksheet.FindSharedStringTablePart();

                        if (sharedStringTablePart != null &&
                            sharedStringTablePart.SharedStringTable != null)
                        {
                            return sharedStringTablePart.SharedStringTable.ElementAt(int.Parse(cell.InnerText)).InnerText;
                        }
                        break;
                    }
                case CellValues.Boolean:
                    {
                        return cell.InnerText == "0" ?
                            bool.FalseString : bool.TrueString;
                    }
            }

            return cell.InnerText;
        }

        public static bool TryParseDate(this Cell cell, out string dateString)
        {
            dateString = null;

            if (cell.StyleIndex == null ||
                !cell.StyleIndex.HasValue)
            {
                return false;
            }

            var worksheet = cell.FindParentWorksheet();
            var document = worksheet.WorksheetPart.OpenXmlPackage as SpreadsheetDocument;
            var styleSheet = document.WorkbookPart.WorkbookStylesPart.Stylesheet;
            var cellStyle = styleSheet.CellFormats.ChildElements[(int)cell.StyleIndex.Value];
            var formatId = (cellStyle as CellFormat).NumberFormatId;

            // See SpreadsheetML Reference at 18.8.30 numFmt(Number Format) for more detail: www.ecma-international.org/publications/standards/Ecma-376.htm
            // Also note some Excel specific variations            
            // The standard defines built-in format ID 14: "mm-dd-yy"; 22: "m/d/yy h:mm"; 37: "#,##0 ;(#,##0)"; 38: "#,##0 ;[Red](#,##0)"; 39: "#,##0.00;(#,##0.00)"; 40: "#,##0.00;[Red](#,##0.00)"; 47: "mmss.0"; KOR fmt 55: "yyyy-mm-dd".
            // Excel defines built-in format ID 14: "m/d/yyyy"; 22: "m/d/yyyy h:mm"; 37: "#,##0_);(#,##0)"; 38: "#,##0_);[Red](#,##0)"; 39: "#,##0.00_);(#,##0.00)"; 40: "#,##0.00_);[Red](#,##0.00)"; 47: "mm:ss.0"; KOR fmt 55: "yyyy/mm/dd".

            if (dateNumberFormats.Contains((int)formatId.Value))
            {
                dateString = DateTime.FromOADate(double.Parse(cell.InnerText)).ToString();
                return true;
            }

            return false;
        }

        public static Worksheet FindParentWorksheet(this Cell cell)
        {
            var parent = cell.Parent;

            while (parent.Parent != null &&
                    parent.Parent != parent &&
                    !parent.LocalName.Equals("worksheet", StringComparison.InvariantCultureIgnoreCase))
            {
                parent = parent.Parent;
            }

            if (!parent.LocalName.Equals("worksheet", StringComparison.InvariantCultureIgnoreCase))
            {
                throw new Exception("Worksheet invalid");
            }

            return parent as Worksheet;
        }

        public static SharedStringTablePart FindSharedStringTablePart(this Worksheet worksheet)
        {
            var document = worksheet.WorksheetPart.OpenXmlPackage as SpreadsheetDocument;

            return document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
        }

        public static string ResolveToNameOrDisplayName(this PropertyInfo item)
        {
            var displayNameAttr = item.GetCustomAttributes(typeof(System.ComponentModel.DisplayNameAttribute), true).Cast<System.ComponentModel.DisplayNameAttribute>().FirstOrDefault();

            if (displayNameAttr != null)
            {
                return displayNameAttr.DisplayName;
            }

            return item.Name;
        }
    }
}
