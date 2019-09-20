---
topic: sample
products:
- office-excel
- office-365
languages:
- csharp
extensions:
  contentType: tools
  createdDate: 6/7/2018 11:27:43 AM
---
# Excel.IO

The goal of this project is to simplify reading and writing Excel workbooks so that the developer needs only pass a collection of objects to write a workbook. Likewise, when reading a workbook the developer supplies a class with properties that map to column names to read a collection of those objects from the workbook. 

Excel.IO takes a single dependency on the [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) and targets .NET Standard 2.0

## Features

* Easy to use developer API
* Write one or more worksheets per workbook by passing a collection of strongly typed objects
* Read one or more worksheets from a workbook into a collection of strongly typed objects

## Limitations

* Assumes workbook structure where the first row has column headers
* Reading multiple worksheets is a little inefficient
* Localisation isn't currently supported 

## Example: Writing a worksheet

Implement [IExcelRow](../master/src/Excel.IO/IExcelRow.cs) and define the columns of the spreadsheet as public properties:

```csharp
public class Person : IExcelRow
{
    public string SheetName { get => "People Sheet"; }

    public string EyeColour { get; set; }

    public int Age { get; set; }

    public int Height { get; set; }
}
```

Then create instances and pass a collection to an instance of [ExcelConverter](../master/src/Excel.IO/ExcelConverter.cs) to write a single sheet workbook with several rows:

```csharp
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
```

## Example: Reading a worksheet

Implement [IExcelRow](../master/src/Excel.IO/IExcelRow.cs) and define the columns of the spreadsheet as public properties (we'll just reuse the same class from above):

```csharp
public class Person : IExcelRow
{
    public string SheetName { get => "People Sheet"; }

    public string EyeColour { get; set; }

    public int Age { get; set; }

    public int Height { get; set; }
}
```

Then, ask an instance of [ExcelConverter](../master/src/Excel.IO/ExcelConverter.cs) to read an IEnumerable<Person> from disk:

```csharp
var excelConverter = new ExcelConverter();
var people = excelConverter.Read<Person>("C:\\somefolder\\people.xlsx");

foreach(var person in people)
{
  //do something useful with the data
}
```

## Feedback

For feature requests or bugs, please [post an issue on GitHub](https://github.com/OfficeDev/Excel-IO/issues).

# Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.microsoft.com.

When you submit a pull request, a CLA-bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., label, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
