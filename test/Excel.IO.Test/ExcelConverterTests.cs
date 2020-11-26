// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using Excel.IO.Test.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Xunit;

namespace Excel.IO.Test
{
	public class ExcelConverterTests : IDisposable
	{
		private Stream xlsxTestResource;

		public ExcelConverterTests()
		{
			var res =
				Assembly.GetExecutingAssembly().GetManifestResourceStream("Excel.IO.Test.Resources.test.xlsx");

			StreamReader sr = new StreamReader(res);
			this.xlsxTestResource = sr.BaseStream;
		}

		public void Dispose()
		{
			this.xlsxTestResource.Close();
			this.xlsxTestResource.Dispose();
		}

		[Fact]
		public void ExcelConverter_Can_Write_A_Single_Sheet_Workbook()
		{
			var excelConverter = new ExcelConverter();

			var sheetName = "Sheet1";
			var rows = new List<IExcelRow>();

			for (int i = 0; i < 100; i++)
			{
				var mockRow = new MockExcelRow();
				mockRow.SheetName = sheetName;

				rows.Add(mockRow);
			}

			using (var result = new MemoryStream())
			{
				excelConverter.Write(rows, result);
				Assert.True(result.Length > 0);
			}
		}

		[Fact]
		public void ExcelConverter_Can_Write_A_MultiSheet_Workbook()
		{
			var excelConverter = new ExcelConverter();
			var rows = new List<IExcelRow>();

			for (int i = 0; i < 100; i++)
			{
				var mockRow = new MockExcelRow();
				mockRow.SheetName = $"Sheet{i}";

				rows.Add(mockRow);
			}

			using (var result = new MemoryStream())
			{
				excelConverter.Write(rows, result);
				Assert.True(result.Length > 0);
			}
		}

		[Fact]
		public void ExcelConverter_Can_Read_A_Single_Sheet_From_A_Workbook()
		{
			var excelConverter = new ExcelConverter();
			var result = excelConverter.Read<MockExcelRow3>(this.xlsxTestResource);

			Assert.Equal(10, result.Count());
		}

		[Fact]
		public void ExcelConverter_Can_Read_Multiple_Sheets_From_A_Workbook()
		{
			var excelConverter = new ExcelConverter();
			var result1 = excelConverter.Read<MockExcelRow3>(this.xlsxTestResource);
			var result2 = excelConverter.Read<MockExcelRow4>(this.xlsxTestResource);
			var result3 = excelConverter.Read<MockExcelRow5>(this.xlsxTestResource);

			Assert.Equal(10, result1.Count());
			Assert.Equal(4, result2.Count());
			Assert.Equal(10, result3.Count());
		}

		[Fact]
		public void ExcelConverter_Can_Read_A_Single_Row_From_A_Sheet()
		{
			var excelConverter = new ExcelConverter();
			List<MockExcelRow5> rows = (List<MockExcelRow5>)excelConverter.Read<MockExcelRow5>(xlsxTestResource);

			Assert.NotEmpty(rows);
			rows[0].GetType().GetProperties().ToList().ForEach(property =>
			{
				Assert.NotNull(property.GetValue(rows[0]));
			});
		}

		[Fact]
		public void ExcelConverter_Can_Read_Multiple_Rows_From_Multiple_Sheet()
		{
			var excelConverter = new ExcelConverter();
			List<MockExcelRow5> rows = (List<MockExcelRow5>)excelConverter.Read<MockExcelRow5>(xlsxTestResource);

			Assert.NotEmpty(rows);

			rows.ForEach(row => {
				row.GetType().GetProperties().ToList().ForEach(property =>
				{
					Assert.NotNull(property.GetValue(row));
				});
			});
		}

		[Fact]
		public void ExcelConverter_Can_Read_Multiple_Rows_From_A_Sheet()
		{
			var excelConverter = new ExcelConverter();
			List<List<IExcelRow>> sheets = new List<List<IExcelRow>>();
			sheets.Add(excelConverter.Read<MockExcelRow3>(this.xlsxTestResource).ToList<IExcelRow>());
			sheets.Add(excelConverter.Read<MockExcelRow4>(this.xlsxTestResource).ToList<IExcelRow>());
			sheets.Add(excelConverter.Read<MockExcelRow5>(this.xlsxTestResource).ToList<IExcelRow>());

			sheets.ForEach(sheet =>
			{
				Assert.NotEmpty(sheet);

				sheet.ForEach(row =>
				{
					row.GetType().GetProperties().ToList().ForEach(property =>
					{
						Assert.NotNull(property.GetValue(row));
					});
				});
			});
		}
	}
}
