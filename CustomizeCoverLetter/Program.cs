using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;


namespace CustomizeCoverLetter
{
    class Program
    {
        static void Main(string[] args)
        {
            string companyname;
            string saveAs;
            bool contains;
            Application word = new Application();


            //open file
            object path = @"F:\CoverLetterSample.docx";

            Document doc = word.Documents.Open(path);

            
            string totaltext = "";

            //ask for company name
            Console.WriteLine("What company are you applying for?: ");
            companyname = Console.ReadLine();

            //search the document for the location
            for (int i = 0; i < doc.Paragraphs.Count; i++)
            {
                contains = doc.Paragraphs[i + 1].Range.Text.ToString().Contains("company");
                if (contains)
                {
                    Console.WriteLine("Yes");

                    //replace user input at the location
                    doc.Paragraphs[i + 1].Range.Text.ToString().Replace("company", companyname);
                    break;
                }
                else
                {
                    Console.WriteLine("No");
                    break;
                }
                //totaltext += " \r\n " + doc.Paragraphs[i + 1].Range.Text.ToString();
            }


            Console.WriteLine("Save document as: ");
            saveAs = Console.ReadLine();
            
            //save new file
            //doc.Save();
            
            doc.Close();
            word.Quit();
        }
    }
}
