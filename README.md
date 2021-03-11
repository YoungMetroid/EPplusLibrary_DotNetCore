# EPPlus_Library_DotNetCore

## Table of Contents
* [Summary](#Summary)
* [Dependencies](#Dependencies)
* [Read Excel Files](#Read-Excel-Files)
* [Create Excel Files](#Create-Excel-Files)

## Summary
This library is a copy of the EPPlus Library but in the .NetCore Framework
This library allows you to read and create excel files.

## Dependencies
To use this library you need to also include the [Utility Library DotNetCore](https://github.azc.ext.hp.com/SLAC-Dev/UtilityLibrary_DotNetCore)

## Read-Excel-Files
* 1-To read from a excel file you'll need to instantiate the Logger since EPPlus uses the logger to catch exceptions.
* 2-Create object of type EPPlusReader and instantiate by passing by parameter the path of where the file is located.
* 3-Set the sheet that you want to read.
* 4-Use the function called `GetTable()` which will get the info thats in the sheet and return a collection of type  `List<List<object>>`.

The reason that a List of List is returned is because its easier to filter information this way since we can filter by rows or by columns using linq.

In the following code snippet you'll see and example on how you can read an excel file and how we filter the information by a specific animal type.

```C#
class Program
	{
		public static Logger logger;
		private const string TestFolderPath = @"C:\TestFolder\";

		static void Main(string[] args)
		{
			logger = Logger.getInstance;
			logger.setLogPathandFile(TestFolderPath, "Error.log");


			EPPlusReader ePPlusReader = new EPPlusReader(TestFolderPath + "Animals.xlsx");
			ePPlusReader.SetSheet(0);
			List<List<object>> table = ePPlusReader.GetTable();

			List<object> dogs = table.Select(x => x[0]).ToList();
			List<object> cats = table.Select(x => x[1]).ToList();
			List<object> birds = table.Select(x => x[2]).ToList();

			dogs.ForEach(x => Console.WriteLine(x));
			Console.WriteLine();
			cats.ForEach(x => Console.WriteLine(x));
			Console.WriteLine();
			birds.ForEach(x => Console.WriteLine(x));
			Console.WriteLine();
			Console.ReadKey();
		}
	}
```

In the following image you'll see the excel file that we are using: 
![alt-text](https://github.azc.ext.hp.com/SLAC-Dev/EPPlus_Library_DotNetCore/blob/master/ReadMe%20Resources/Example2.PNG)

In the following image you'll see the results you should get when you run the run:
![alt-text](https://github.azc.ext.hp.com/SLAC-Dev/EPPlus_Library_DotNetCore/blob/master/ReadMe%20Resources/Example3.PNG)

## Create Excel Files

To create a excel file and save info to it you'll need to follow these steps:
* 1-Create a `EPPlusCreator()` object and instantiate it.
* 2-Create a Sheet and set the `EPPlusCreator` object to that sheet.
* 3-Use the `WriteInfo()` function and pass by parameter a Collection of type <List<List<object>> .
* 4-Use the `SaveFile()` function and pass by parameter the path of where you want the file to be saved at and a boolean, false meaning you want the file to be saved as `xslx` or true  to save the file as `xlsm`.
	
In the following code snippet you'll see an example on how you can create a excel file and write info to it:

```C#
	class Program
	{
		public static Logger logger;
		private const string TestFolderPath = @"C:\TestFolder\";

		static void Main(string[] args)
		{
			logger = Logger.getInstance;
			logger.setLogPathandFile(TestFolderPath, "Error.log");


			List<List<object>> names = new List<List<object>>();
			List<object> nameInfo = new List<object>();


			nameInfo.AddRange(new List<object> { "FirstName", "LastName" });
			names.AddRange(new List<List<object>> { nameInfo });

			nameInfo = new List<object>();
			nameInfo.AddRange(new List<object> { "Bob", "Jones1" });
			names.AddRange(new List<List<object>> { nameInfo });

			nameInfo = new List<object>();
			nameInfo.AddRange(new List<object> { "Phillip", "Jones2" });
			names.AddRange(new List<List<object>> { nameInfo });


			nameInfo = new List<object>();
			nameInfo.AddRange(new List<object> { "Mine", "Jones3" });
			names.AddRange(new List<List<object>> { nameInfo });

			nameInfo = new List<object>();
			nameInfo.AddRange(new List<object> { "Craft", "Jones4" });
			names.AddRange(new List<List<object>> { nameInfo });

			EPPlusCreator ePPlus = new EPPlusCreator();
			ePPlus.CreateSheet("TestSheet");
			ePPlus.SetSheet("TestSheet");
			ePPlus.WriteInfo(names);
			ePPlus.SaveFile(TestFolderPath + "EPPlus.xlsx", false);
		}
	}
```

When you run the code the `EPPlus.xlsx` file should be created and it should have the following info:
![alt-text](https://github.azc.ext.hp.com/SLAC-Dev/EPPlus_Library_DotNetCore/blob/master/ReadMe%20Resources/Example1.PNG)
