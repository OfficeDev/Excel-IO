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
    public class ExcelConverterTestsTwo : IDisposable
    {
        private Stream xlsxTestResource;

        public ExcelConverterTestsTwo()
        {
            var res =
                Assembly.GetExecutingAssembly().GetManifestResourceStream("Excel.IO.Test.Resources.test2.xlsx");

            StreamReader sr = new StreamReader(res);
            this.xlsxTestResource = sr.BaseStream;
        }

        [Fact]
        public void ExcelConverter_Can_Read_A_Single_Sheet_Workbook()
        {
            var excelConverter = new ExcelConverter();
            List<MockExcelRowTwo> rows = (List<MockExcelRowTwo>)excelConverter.Read<MockExcelRowTwo>(xlsxTestResource);

            Assert.NotEmpty(rows);
            rows[0].GetType().GetProperties().ToList().ForEach(property =>
            {
                Assert.NotNull(property.GetValue(rows[0]));
            });
        }

        public void Dispose()
        {
            this.xlsxTestResource.Close();
            this.xlsxTestResource.Dispose();
        }
    }
}
