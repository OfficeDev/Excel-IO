﻿// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Excel.IO
{
    /// <summary>
    /// Converter that allows implementations of <see cref="IExcelRow "/> to be exported.
    /// </summary>
    public class ExcelConverter : IExcelConverter
    {
        private SpreadsheetDocument _GetDocument(Stream stream)
        {
            var spreadsheetDocument = SpreadsheetDocument.Open(stream, isEditable: true);

            if (spreadsheetDocument.WorkbookPart == null)
            {
                return SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            }

            return spreadsheetDocument;            
        }

        public void Append(IExcelRow row, Stream outputStream)
        {
            using (var spreadsheetDocument = _GetDocument(outputStream))
            {
                this.Write([row], spreadsheetDocument);
            }
        }

        /// <summary>
        /// Exports the given rows to an Excel workbook
        /// </summary>
        /// <param name="rows">The rows to write to the workbook. Each property will be written as a cell in the row.</param>
        /// <param name="outputStream">The stream to write the workbook to</param>
        public void Write(IEnumerable<IExcelRow> rows, Stream outputStream)
        {
            using (var spreadsheetDocument = SpreadsheetDocument.Create(outputStream, SpreadsheetDocumentType.Workbook))
            {
                this.Write(rows, spreadsheetDocument);
            }
        }

        /// <summary>
        /// Exports the given rows to an Excel workbook
        /// </summary>
        /// <param name="rows">The rows to write to the workbook. Each property will be written as a cell in the row.</param>
        /// <param name="path">The path to write the workbook to</param>
        public void Write(IEnumerable<IExcelRow> rows, string path)
        {
            using (var spreadsheetDocument = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
            {
                this.Write(rows, spreadsheetDocument);
            }
        }

        private void Write(IEnumerable<IExcelRow> rows, SpreadsheetDocument spreadsheetDocument)
        {
            if (spreadsheetDocument.WorkbookPart == null)
            {
                var workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();
            }

            var sheets = spreadsheetDocument.WorkbookPart.Workbook.Sheets;

            if (sheets == null)
            {
                sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
            }

            var rowsGroupedBySheet = rows.GroupBy(r => r.SheetName);

            uint sheetId = 1;

            foreach (var rowGroup in rowsGroupedBySheet)
            {
                var sheetData = default(SheetData);
                var headerWritten = false;
                uint rowIndex = 1;

                var existingSheet = sheets.ChildElements.OfType<Sheet>().FirstOrDefault(s => s.Name == rowGroup.Key);

                if (existingSheet == null)
                {
                    var worksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());

                    var relationshipIdPart = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart);
                    var sheet = new Sheet() { Id = relationshipIdPart, SheetId = sheetId, Name = rowGroup.Key };

                    sheets.Append(sheet);
                    sheetId++;

                    sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                }
                else
                {
                    var worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(existingSheet.Id);
                    sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();                                       

                    // get the correct row to write to
                    var lastSheetRow = sheetData.ChildElements.OfType<Row>().Last();
                    rowIndex = lastSheetRow.RowIndex + 1;
                    headerWritten = true;
                }

                // write rows to this sheet
                var propertiesToIgnore = typeof(IExcelRow).GetProperties();                                

                foreach (var row in rowGroup)
                {
                    var sheetRow = new Row { RowIndex = new UInt32Value(rowIndex) };
                    sheetData.Append(sheetRow);

                    var properties = row.GetType().GetProperties();
                    var validProperties = properties.Except(propertiesToIgnore, SimpleComparer.Instance);

                    if (!headerWritten)
                    {
                        this.WriteHeader(validProperties, sheetRow, row);

                        headerWritten = true;
                        rowIndex++;

                        sheetRow = new Row { RowIndex = new UInt32Value(rowIndex) };
                        sheetData.Append(sheetRow);
                    }

                    this.WriteCells(validProperties, sheetRow, row);
                    
                    rowIndex++;
                }
            }
        }

        /// <summary>
        /// Reads a known workbook format into a collection of IExcelRow implementations
        /// </summary>
        /// <typeparam name="T">Implementation of IExcelRow that specifies the sheet to read and the row headings to include</typeparam>
        /// <param name="path">Path on disk of the workbook</param>
        /// <returns>A collection of <typeparamref name="T"/></returns>
        public IEnumerable<T> Read<T>(string path) where T : IExcelRow, new()
        {
            using (var spreadsheetDocument = SpreadsheetDocument.Open(path, false))
            {
                return this.Read<T>(spreadsheetDocument);
            }
        }

        /// <summary>
        /// Reads a known workbook format into a collection of IExcelRow implementations
        /// </summary>
        /// <typeparam name="T">Implementation of IExcelRow that specifies the sheet to read and the row headings to include</typeparam>
        /// <param name="stream">Stream that represents the workbook</param>
        /// <returns>A collection of <typeparamref name="T"/></returns>
        public IEnumerable<T> Read<T>(Stream stream) where T : IExcelRow, new()
        {
            using (var spreadsheetDocument = SpreadsheetDocument.Open(stream, false))
            {
                return this.Read<T>(spreadsheetDocument);
            }
        }

        private IEnumerable<T> Read<T>(SpreadsheetDocument spreadsheetDocument) where T : IExcelRow, new()
        {
            var toReturn = new List<T>();
            var workBookPart = spreadsheetDocument.WorkbookPart;

            foreach (var sheet in workBookPart.Workbook.Descendants<Sheet>())
            {
                var worksheetPart = workBookPart.GetPartById(sheet.Id) as WorksheetPart;

                if (worksheetPart == null)
                {
                    // the part was supposed to be here, but wasn't found :/
                    continue;
                }

                if (sheet.Name.HasValue && sheet.Name.Value.Equals(new T().SheetName))
                {
                    toReturn.AddRange(this.ReadSheet<T>(worksheetPart));
                }
            }

            return toReturn;
        }

        private List<T> ReadSheet<T>(WorksheetPart wsPart) where T : IExcelRow, new()
        {
            var toReturn = new List<T>();

            // assume the first row contains column names
            var headerRow = true;
            var headers = new Dictionary<string, object>();

            foreach (var row in wsPart.Worksheet.Descendants<Row>())
            {
                // one instance of T per row
                var obj = new T();
                var properties = obj.GetType().GetProperties();

                foreach (Cell c in row.Elements<Cell>())
                {
                    var column = c.GetColumn();
                    var value = c.GetCellValue();

                    if (headerRow)
                    {
                        headers.Add(column, value);
                    }
                    else
                    {
                        // look for a property on the T that matches the name (ignore SheetName)
                        object columnHeader = null;

                        if (headers.TryGetValue(column, out columnHeader))
                        {
                            var propertyInfo = properties.Where(p =>
                                p.ResolveToNameOrDisplayName().Equals(columnHeader.ToString(), StringComparison.OrdinalIgnoreCase)).FirstOrDefault();

                            if (propertyInfo != null)
                            {
                                Type t = propertyInfo.PropertyType;
                                t = Nullable.GetUnderlyingType(t) ?? t;

                                if (t.IsEnum)
                                {
                                    propertyInfo.SetValue(obj, Enum.Parse(t, (string)value));
                                }
                                else
                                {
                                    propertyInfo.SetValue(obj, Convert.ChangeType(value, t));
                                }
                            }
                        }
                    }
                }

                if (!headerRow)
                {
                    toReturn.Add(obj);
                }

                headerRow = false;
            }

            return toReturn;
        }

        private void WriteCells(IEnumerable<PropertyInfo> properties, Row sheetRow, IExcelRow userRow)
        {
            var columnIndex = 0;

            foreach (var item in properties)
            {
                var result = _TryInsertExcelColumn(sheetRow, userRow, columnIndex, item, isHeader: false);

                if (result.IsExcelColumn)
                {
                    columnIndex = result.ColumnIndex;
                    continue;
                }

                var cellValue = item.GetValue(userRow);

                sheetRow.InsertAt(
                    new Cell
                    {
                        CellReference = sheetRow.GetCellReference(columnIndex + 1),
                        CellValue = new CellValue(cellValue == null ? string.Empty : cellValue.ToString()),
                        DataType = new EnumValue<CellValues>(this.ResolveCellType(item.PropertyType))
                    },
                    columnIndex);

                columnIndex++;
            }
        }

        private CellValues ResolveCellType(Type propertyType)
        {
            var nullableType = Nullable.GetUnderlyingType(propertyType);

            if (nullableType != null)
            {
                propertyType = Nullable.GetUnderlyingType(propertyType);
            }

            // TODO: Support date? 
            switch (Type.GetTypeCode(propertyType))
            {
                case TypeCode.Decimal:
                case TypeCode.Double:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                    {
                        return CellValues.Number;
                    }
                default:
                    {
                        return CellValues.String;
                    }
            }
        }

        private void WriteHeader(IEnumerable<PropertyInfo> properties, Row sheetRow, IExcelRow userRow)
        {
            var columnIndex = 0;

            foreach (var item in properties)
            {
                var result = _TryInsertExcelColumn(sheetRow, userRow, columnIndex, item, isHeader: true);

                if (result.IsExcelColumn)
                {
                    columnIndex = result.ColumnIndex;
                    continue;
                }

                var headerName = item.Name;

                var displayNameAttr = item.GetCustomAttribute<System.ComponentModel.DisplayNameAttribute>(true);

                if (displayNameAttr != null)
                {
                    headerName = displayNameAttr.DisplayName;
                }

                sheetRow.InsertAt(
                    new Cell
                    {
                        CellReference = sheetRow.GetCellReference(columnIndex + 1),
                        CellValue = new CellValue(headerName),
                        DataType = new EnumValue<CellValues>(CellValues.String)
                    },
                    columnIndex);

                columnIndex++;
            }
        }

        private InsertExcelColumnResult _TryInsertExcelColumn(Row sheetRow, IExcelRow row, int columnIndex, PropertyInfo item, bool isHeader)
        {
            var excelColumnsAttr = item.GetCustomAttribute<ExcelColumnsAttribute>(true);

            if (excelColumnsAttr != null)
            {
                var dict = (IDictionary<string, string>)item.GetValue(row);

                if (dict != null)
                {
                    foreach (var kvp in dict)
                    {
                        sheetRow.InsertAt(
                            new Cell
                            {
                                CellReference = sheetRow.GetCellReference(columnIndex + 1),
                                CellValue = new CellValue(isHeader ?
                                    kvp.Key :
                                        kvp.Value == null ?
                                            string.Empty : kvp.Value),
                                DataType = new EnumValue<CellValues>(isHeader ?
                                    CellValues.String :
                                    this.ResolveCellType(item.PropertyType))
                            },
                            columnIndex);

                        columnIndex++;
                    }

                    return new InsertExcelColumnResult { IsExcelColumn = true, ColumnIndex = columnIndex };
                }
            }

            return InsertExcelColumnResult.IsNotExcelColumn;
        }

        private class InsertExcelColumnResult
        {
            private static readonly InsertExcelColumnResult _IsNotExcelColumn = new InsertExcelColumnResult { IsExcelColumn = false };

            public static InsertExcelColumnResult IsNotExcelColumn
            {
                get { return _IsNotExcelColumn; }
            }

            public int ColumnIndex { get; set; }

            public bool IsExcelColumn { get; set; }
        }

        private class SimpleComparer : IEqualityComparer<PropertyInfo>
        {
            private static readonly SimpleComparer ReadonlyInstance;

            static SimpleComparer()
            {
                ReadonlyInstance = new SimpleComparer();
            }

            public static SimpleComparer Instance
            {
                get { return ReadonlyInstance; }
            }

            public bool Equals(PropertyInfo x, PropertyInfo y)
            {
                return x.Name == y.Name;
            }

            public int GetHashCode(PropertyInfo obj)
            {
                // only care if the name of the property info matches
                return obj.Name.GetHashCode();
            }
        }
    }
}
