using Excel.IO;
using Excel.IO.examples;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExampleReadAndWrite
{
    class Program
    {
        static void Main(string[] args)
        {
            var excelConverter = new ExcelConverter();
            var cwd = Directory.GetCurrentDirectory();
            Console.WriteLine(cwd);
            //Read Example            
            var people = excelConverter.Read<Person>("..\\..\\..\\people.xlsx");

            foreach (var person in people)
            {
                Console.WriteLine($"{person.EyeColour} : {person.Age} : {person.Height}");
            }

            //Write Example
            var peopleToWrite = new List<Person>();

            for (int i = 0; i < 10; i++)
            {
                peopleToWrite.Add(new Person
                {
                    EyeColour = Guid.NewGuid().ToString(),
                    Age = new Random().Next(1, 100),
                    Height = new Random().Next(100, 200)
                });
            }
            
            excelConverter.Write(peopleToWrite, "..\\..\\..\\newPeople.xlsx");
        }
    }
}

