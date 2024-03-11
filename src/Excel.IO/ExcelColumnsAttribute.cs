// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System;

namespace Excel.IO
{
    /// <summary>
    /// An attribute that allows non-property fields to be used as columns in an Excel file. This is only intended for use with IDictionary<string, string>.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelColumnsAttribute : Attribute
    { }
}
