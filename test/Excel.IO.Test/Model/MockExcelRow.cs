// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System;

namespace Excel.IO.Test.Model
{
    public class MockExcelRow : IExcelRow
    {
        public string SheetName { get; set; }

        public DateTime LastContact { get; set; }

        public int CustomerId { get; set; }

        public string FirstName { get; set; }

        public string LastName { get; set; }

        public string Address { get; set; }

        public bool IsActive { get; set; }

        public decimal Balance { get; set; }

        public Category Category { get; set; }
    }

    public class MockExcelRow2 : IExcelRow
    {
        public string SheetName { get; set; }

        public string EyeColour { get; set; }

        public int Age { get; set; }

        public int Height { get; set; }
    }

    public class MockExcelRow3 : MockExcelRow
    {
        public MockExcelRow3()
        {
            this.SheetName = "Sheet1";
        }
    }

    public class MockExcelRow4 : MockExcelRow2
    {
        public MockExcelRow4()
        {
            this.SheetName = "Sheet2";
        }
    }

    public enum Category
    {
        CategoryA,
        CategoryB,
        CategoryC
    }
}
