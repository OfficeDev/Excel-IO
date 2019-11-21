using Excel.IO;
using System;
using System.Collections.Generic;

namespace Excel.IO.examples
{
    class WriteExample
    {
        static void Main(string[] args)
        {                          
            var people = new List<Person>();
            
            for (int i = 0; i < 10; i++)
            {
                people.Add(new Person
                {
                    EyeColour = Guid.NewGuid().ToString(),
                    Age = new Random().Next(1, 100),
                    Height = new Random().Next(100, 200)
                });
            }
            var excelConverter = new ExcelConverter();
            excelConverter.Write(people, "C:\\somefolder\\people.xlsx");                      
        }
    }  
}
