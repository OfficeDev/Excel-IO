using Excel.IO;
using System;
using System.Collections.Generic;

namespace Excel.IO.examples
{
    class ReadExample
    {
        static void notmain(string[] args)
        {            
            var excelConverter = new ExcelConverter();
            var people = excelConverter.Read<Person>("C:\\somefolder\\people.xlsx");

            foreach (var person in people)
            {
                Console.WriteLine($"{person.EyeColour} : {person.Age} : {person.Height}");
            }
            
        }
    }   
}
