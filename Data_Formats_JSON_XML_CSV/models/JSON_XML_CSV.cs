using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using JsonSerializer = System.Text.Json.JsonSerializer;

namespace Data_Formats_JSON_XML_CSV.models
{
    internal class JSON_XML_CSV
    {
        string Categories = "Category";
        string Titles = "Title";
        string Authors = "Author";
        string Years = "Year";
        string Prices = "Price";


        public class Book
        {
            public string Title { get; set; }
            public string[] Authors { get; set; }
            public int Year { get; set; }
            public decimal Price { get; set; }
            public string Category { get; set; }
            public string Cover { get; set; }
        }

        public class Bookstore
        {
            public Book[] Books { get; set; }
        }

        
            public Bookstore JsonTableCreator(string jsonFilePath)
            {
                string jsonData = System.IO.File.ReadAllText(jsonFilePath);
                JObject jsonObject = JObject.Parse(jsonData);

                JArray bookArray = jsonObject["bookstore"]["book"] as JArray;
                if (bookArray == null)
                    throw new ArgumentException("Invalid JSON format");

                Bookstore bookstore = new Bookstore();
                bookstore.Books = new Book[bookArray.Count];

                for (int i = 0; i < bookArray.Count; i++)
                {
                    JObject bookObject = bookArray[i] as JObject;
                    if (bookObject == null)
                        throw new ArgumentException("Invalid JSON format");

                    bookstore.Books[i] = new Book
                    {
                        Title = bookObject["title"]["__text"].ToString(),
                        Authors = bookObject["author"].Type == JTokenType.Array
                            ? bookObject["author"].ToObject<string[]>()
                            : new string[] { bookObject["author"].ToString() },
                        Year = int.Parse(bookObject["year"].ToString()),
                        Price = decimal.Parse(bookObject["price"].ToString()),
                        Category = bookObject["_category"].ToString(),
                        Cover = bookObject["_cover"]?.ToString()
                    };
                }

                return bookstore;
            }
        
       
        public void XmlTableCreator(string xmlPath)
        {
            
            // Create an XML document and load the XML file
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlPath);

            // Get the list of book nodes
            XmlNodeList bookNodes = xmlDoc.SelectNodes("/books/book");

            // Create a table
            Console.WriteLine("-----------------------------------------------------------------------");
            Console.WriteLine($"| {Categories.PadRight(25)} | {Titles.PadRight(25)} | {Authors.PadRight(70)} |  {Prices} | ");
            Console.WriteLine("-----------------------------------------------------------------------");

            foreach (XmlNode bookNode in bookNodes)
            {
                string title = bookNode.SelectSingleNode("title").InnerText;
                string authors = "";
                XmlNodeList authorNodes = bookNode.SelectNodes("author");
                foreach (XmlNode authorNode in authorNodes)
                {
                    authors += authorNode.InnerText + ", ";
                }
                authors = authors.TrimEnd(',', ' ');

                string year = bookNode.SelectSingleNode("year").InnerText;
                string price = bookNode.SelectSingleNode("price").InnerText;
                string category = bookNode.Attributes["category"].Value;
                string cover = bookNode.Attributes["cover"]?.Value ?? "";

                Console.WriteLine($"| {category.PadRight(25)} | {title.PadRight(25)} | {authors.PadRight(70)} |  {price} | ");
            }

            Console.WriteLine("-----------------------------------------------------------------------");
        }

        
            public void CsvTableCreator(string csvPath)
            {
                // Read all lines from the CSV file
                string[] lines = System.IO.File.ReadAllLines(csvPath);

                // Create a table
                Console.WriteLine("-----------------------------------------------------------");
                 Console.WriteLine($"| {Categories.PadRight(25)} | {Titles.PadRight(25)} | {Authors.PadRight(70)} |  {Prices} | ");
                Console.WriteLine("-----------------------------------------------------------");

                // Skip the header line (first line) and process each data line
                for (int i = 1; i < lines.Length; i++)
                {
                    string[] data = lines[i].Split(',');

                    // Ensure that data array has at least 5 elements (Category, Title, Author, Year, Price)
                    if (data.Length >= 5)
                    {
                        string category = data[0].Trim();
                        string title = data[1].Trim();
                        string author = data[2].Trim();
                        string year = data[3].Trim();
                        string price = data[4].Trim();

                    Console.WriteLine($"| {category.PadRight(25)} | {title.PadRight(25)} | {author.PadRight(70)} |  {price} | ");
                }
                }

                Console.WriteLine("-----------------------------------------------------------");
            }
        public void CreateExcelFile(string jsonFilePath, string excelFilePath)
        {
            Bookstore bookstore = JsonTableCreator(jsonFilePath);

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Books");

                // Write header row
                worksheet.Cells[1, 1].Value = "Title";
                worksheet.Cells[1, 2].Value = "Author(s)";
                worksheet.Cells[1, 3].Value = "Year";
                worksheet.Cells[1, 4].Value = "Price";
                worksheet.Cells[1, 5].Value = "Category";
                worksheet.Cells[1, 6].Value = "Cover";

                // Write data rows
                int rowIndex = 2;
                foreach (var book in bookstore.Books)
                {
                    worksheet.Cells[rowIndex, 1].Value = book.Title;
                    worksheet.Cells[rowIndex, 2].Value = string.Join(", ", book.Authors);
                    worksheet.Cells[rowIndex, 3].Value = book.Year;
                    worksheet.Cells[rowIndex, 4].Value = book.Price;
                    worksheet.Cells[rowIndex, 5].Value = book.Category;
                    worksheet.Cells[rowIndex, 6].Value = book.Cover ?? "Unknown";

                    rowIndex++;
                }

                package.SaveAs(new FileInfo(excelFilePath));
            }
        }

    }
}
