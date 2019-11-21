using System;
using System.Collections.Generic;
using System.Text;

namespace Excel.IO.examples
{
    public class Person : IExcelRow
    {
        public string SheetName { get => "People Sheet"; }

        public string EyeColour { get; set; }

        public int Age { get; set; }

        public int Height { get; set; }
    }
}
