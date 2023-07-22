
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System.Diagnostics;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Numerics;
using System.Reflection.PortableExecutable;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using static System.Runtime.InteropServices.JavaScript.JSType;


namespace Data_Formats_JSON_XML_CSV.models
{
    internal class JsonDataFetcher
    {
        public async Task<string> GetJsonFromUrlAsync(string url)
        {
            using (var httpClient = new HttpClient())
            {
                try
                {
                    HttpResponseMessage response = await httpClient.GetAsync(url);
                    response.EnsureSuccessStatusCode();
                    return await response.Content.ReadAsStringAsync();
                }
                catch (HttpRequestException e)
                {
                    Console.WriteLine($"Error while fetching data from URL: {e.Message}");
                    return string.Empty;
                }
            }
        }

        public void DisplayUserName(JArray jsonArray)
        {
            Console.WriteLine("Name");

            foreach (JObject user in jsonArray)
            {
                string name = user["name"].ToString();

                Console.WriteLine($"{name}");
            }
        }
        public void DisplayUserEmail(JArray jsonArray)
        {
            Console.WriteLine("Email");

            foreach (JObject user in jsonArray)
            {
                string email = user["email"].ToString();

                Console.WriteLine($"{email}");

            }
        }
        public void DisplayUserPhone(JArray jsonArray)
        {
            Console.WriteLine("Phone");

            foreach (JObject user in jsonArray)
            {
                string phone = user["phone"].ToString();

                Console.WriteLine($"{phone}");

            }
        }
        public void DisplayUserAddress(JArray jsonArray)
        {
            Console.WriteLine("Phone");

            foreach (JObject user in jsonArray)
            {
                string address = $"{user["address"]["suite"].ToString()}-{user["address"]["street"].ToString()}, {user["address"]["city"].ToString()}";

                Console.WriteLine($"{address}");


            }
        }
        //static void DisplayUserInfoTable(JArray jsonArray)
        //{
        //    Console.WriteLine("Name\t\tEmail\t\t\t\tPhone\t\t\tAddress\n");

        //    foreach (JObject user in jsonArray)
        //    {
        //        string name = user["name"].ToString();
        //        string email = user["email"].ToString();
        //        string phone = user["phone"].ToString();
        //        string address = $"{user["address"]["suite"].ToString()}-{user["address"]["street"].ToString()}, {user["address"]["city"].ToString()}";

        //        Console.WriteLine($"{name}\t\t{email}\t\t{phone}\t\t{address}\n");
        //    }
        //}

        public void DisplayInformationInTable(string jsonData)
        {
            List<User> users = JsonConvert.DeserializeObject<List<User>>(jsonData);

            StringBuilder tableBuilder = new StringBuilder();
            tableBuilder.AppendLine("Name\t\t\t\t Email\t\t\t\t Phone\t\t\t\t Address");
            tableBuilder.AppendLine("======================================================================");
            foreach (var user in users)
            {
                string name = user.name;
                string email = user.email;
                string phone = user.phone;
                string address = $"{user.address.suite}-{user.address.street}, {user.address.city}";
                tableBuilder.AppendLine($"{name.PadRight(30)} {email.PadRight(30)}{phone.PadRight(20)}{address.PadRight(50)}");
            }

            Console.WriteLine(tableBuilder.ToString());
        }

        class User
        {
            public string name { get; set; }
            public string email { get; set; }
            public string phone { get; set; }
            public Address address { get; set; }
        }

        class Address
        {
            public string street { get; set; }
            public string suite { get; set; }
            public string city { get; set; }
            // Add other address fields if needed (e.g., zipcode, geo, etc.)
        }
        public void SaveToExcel(JArray jsonArray)
        {
            string excelFileName = "Users.xlsx";

            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Users");

                worksheet.Cells[1, 1].Value = "Name";
                worksheet.Cells[1, 2].Value = "Email";
                worksheet.Cells[1, 3].Value = "Phone";
                worksheet.Cells[1, 4].Value = "Address";

                int row = 2;

                foreach (JObject user in jsonArray)
                {
                    string name = user["name"].ToString();
                    string email = user["email"].ToString();
                    string phone = user["phone"].ToString();
                    string address = $"{user["address"]["suite"].ToString()}-{user["address"]["street"].ToString()}, {user["address"]["city"].ToString()}";

                    worksheet.Cells[row, 1].Value = name;
                    worksheet.Cells[row, 2].Value = email;
                    worksheet.Cells[row, 3].Value = phone;
                    worksheet.Cells[row, 4].Value = address;

                    row++;
                }

                FileInfo excelFile = new FileInfo(excelFileName);
                excelPackage.SaveAs(excelFile);

                Console.WriteLine($"Data saved to {excelFileName}");

                // Get the full path of the created Excel file
                string excelFilePath = Path.Combine(Directory.GetCurrentDirectory(), excelFileName);

                Console.WriteLine();
                Console.WriteLine("\n ============ Show file location ============== \n");
                Console.WriteLine();
                Console.WriteLine($"Data saved to {excelFilePath}");


                // Open the Excel file with the default associated application (Microsoft Excel)
                try
                {
                    // Check if the operating system is Windows
                    if (Environment.OSVersion.Platform == PlatformID.Win32NT)
                    {
                        Process.Start(excelFile.FullName);
                    }
                    else
                    {
                        Console.WriteLine("Opening Excel files automatically is not supported on this operating system.");
                        Console.WriteLine($"Please open the file manually: {excelFile.FullName}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error while trying to open the Excel file: {ex.Message}");
                    Console.WriteLine($"Please open the file manually: {excelFile.FullName}");
                }
            }
        }

    }
}

