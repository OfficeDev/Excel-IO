﻿// Copyright (c) Microsoft Corporation. All rights reserved.
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

    public class MockExcelRow5 : IExcelRow
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

        public int Age { get; set; }

        public bool IsMarried { get; set; }

        public string PhoneNumber { get; set; }

        public string Email { get; set; }

        public decimal Debt { get; set; }

        public decimal HouseholdIncome { get; set; }

        public float AgePercentage { get; set; }

        public DateTime BirthDate { get; set; }

        public float ProbabilityOfSameAge { get; set; }

        public float Constants { get; set; }

        public DateTime LongDate { get; set; }

        public DateTime LongDate2 { get; set; }

        public DateTime DayMonth { get; set; }

        public DateTime Something { get; set; }

        public Category Category1 { get; set; }

        public Category Category2 { get; set; }

        public Category Category3 { get; set; }

        public Category Category4 { get; set; }

        public Category Category5 { get; set; }

        public Category Category6 { get; set; }

        public Category Category7 { get; set; }
public MockExcelRow5() { SheetName = "Sheet3"; }
    }

    public enum Category
    {
        CategoryA,
        CategoryB,
        CategoryC
    }
}
