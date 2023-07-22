using Data_Formats_JSON_XML_CSV.models;
using Newtonsoft.Json.Linq;
using System.Diagnostics;
using static Data_Formats_JSON_XML_CSV.models.JSON_XML_CSV;

string url = "https://jsonplaceholder.typicode.com/users";
JsonDataFetcher jsonDataFetcher = new();
string jsonData = await jsonDataFetcher.GetJsonFromUrlAsync(url);

// Display the JSON data in the specified format

Console.WriteLine("\n ============ Read Data From URL ============== \n");
Console.WriteLine();
Console.WriteLine(jsonData);
JArray jsonArray = JArray.Parse(jsonData);
Console.WriteLine("\n ============ List Of User Name ============== \n");
Console.WriteLine();
jsonDataFetcher.DisplayUserName(jsonArray);
Console.WriteLine("\n ============ List Of User Email ============== \n");
Console.WriteLine();
jsonDataFetcher.DisplayUserEmail(jsonArray);
Console.WriteLine("\n ============ List Of User Phone Number ============== \n");
Console.WriteLine();
jsonDataFetcher.DisplayUserPhone(jsonArray);
Console.WriteLine("\n ============ List Of User Address ============== \n");
Console.WriteLine();
jsonDataFetcher.DisplayUserAddress(jsonArray);
Console.WriteLine("\n ============ Table for User Information ============== \n");
Console.WriteLine();
jsonDataFetcher.DisplayInformationInTable(jsonData);
//Console.WriteLine("\n ============ Export User Information ============== \n");
//Console.WriteLine();
//jsonDataFetcher.SaveToExcel(jsonArray);



string csvFilePath = @"C:\Users\akopr\OneDrive\Documents\Books\books.csv";
string xmlFilePath = @"C:\Users\akopr\OneDrive\Documents\Books\books.xml";
string jsonFilePath = @"C:\Users\akopr\OneDrive\Documents\Books\books.json";
//XmlTableCreator.DisplayBooks(xmlFilePath);
JSON_XML_CSV jsonXmlCsv = new();
//var books = ReadBooksFromJson(jsonFilePath);
//DisplayTable(books);
Console.WriteLine("\n ============ Table for XMl file ============== \n");
Console.WriteLine();
jsonXmlCsv.XmlTableCreator(xmlFilePath);
Console.WriteLine("\n ============ Table for CSV file ============== \n");
Console.WriteLine();
jsonXmlCsv.CsvTableCreator(csvFilePath);
Console.WriteLine("\n ============ Table for JSON file ============== \n");
Console.WriteLine();

Bookstore bookstore = jsonXmlCsv.JsonTableCreator(jsonFilePath);

// Now you can access the books in the bookstore object and display or process them as needed.
foreach (var book in bookstore.Books)
{
    
    Console.WriteLine($"| {book.Category.PadRight(25)} | {book.Title.PadRight(25)} | {string.Join(", ", book.Authors).PadRight(70)} |  {book.Price} | ");
}
    Console.WriteLine();
    //jsonXmlCsv.JsonTableCreator(jsonFilePath);

Console.WriteLine("\n ============ Excel file for JSON file ============== \n");
Console.WriteLine();
string excelFilePath = "Bookstore.xlsx";

//jsonXmlCsv.CreateExcelFile(jsonFilePath, excelFilePath);
Console.WriteLine($"Excel file created: {excelFilePath}");
