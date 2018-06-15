// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System.Collections.Generic;
using System.IO;

namespace Excel.IO
{
    public interface IExcelConverter
    {
        void Write(IEnumerable<IExcelRow> rows, Stream outputStream);

        void Write(IEnumerable<IExcelRow> rows, string path);

        IEnumerable<T> Read<T>(string path) where T : IExcelRow, new();

        IEnumerable<T> Read<T>(Stream stream) where T : IExcelRow, new();
    }
}
