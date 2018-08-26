using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace CsExcelToXml
{
    class Program
    {
        const string FILEPATH_INPUT = "input.xlsx";
        const string FILEPATH_OUTPUT = "output.xml";

        static void Main(string[] args)
        {
            Console.WriteLine("Opening Excel File...");
            XSSFWorkbook workbook;
            using (FileStream file = new FileStream(FILEPATH_INPUT, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(file);
            }

            Console.WriteLine("Selecting first sheet...");
            ISheet sheet = workbook.GetSheetAt(0); // get first sheet
            
            Console.WriteLine("\nReading Excel Rows into C# list...");
            List<Student> students = new List<Student>();
            for (int row = 1; row <= sheet.LastRowNum; row++) // Started from row 1 because row 0 is header
            {
                Console.Write($"{row}...");
                Student student = new Student();

                student.Id = sheet.GetRow(row).GetCell(0).ToString().Trim();
                student.Name = sheet.GetRow(row).GetCell(1).ToString().Trim();
                student.Family = sheet.GetRow(row).GetCell(2).ToString().Trim();

                students.Add(student);
            }
            Console.WriteLine("Reading completed.");

            Console.WriteLine("\nWriting C# list into XML file...");
            XmlSerializer writer = new XmlSerializer(typeof(List<Student>));
            using (FileStream stream = File.Create(FILEPATH_OUTPUT))
            {
                writer.Serialize(stream, students);
            }
            Console.WriteLine("Writing completed.");

            Console.WriteLine("\nPress any key to exit...");
            Console.ReadKey();
        }
    }
}
