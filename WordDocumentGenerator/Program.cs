using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Spire.Doc;
using Spire.Doc.Documents;
using WordDocumentGenerator.Models;

namespace WordDocumentGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            setConsoleTitle();

            //Create new Document
            Document BlankDocument = new Document();

            // Cargar un Documento existente
            Document document = new Document();
            document.LoadFromFile(@"D:\Credentials\GoogleDocumentGenerated.docx");

            //Generar nuevo documento word
            GenerateDocument(BlankDocument);

            // contar cuantos parrafos tiene una seccion
            //CountParagraph(document);

            Console.ReadKey();
        }

        public static void GenerateDocument(Document document)
        {
            Console.WriteLine(">>>>> Generation of Word Document <<<<<\n");

            // Get data of the API
            var audit = GetAudit().Result;

            // Add Section
            Section section = document.AddSection();

            //----------------------- TITLE --------------------
            // Add new paragraph
            Paragraph pTitle = section.AddParagraph();
            // Add Title
            pTitle.AppendText(audit.Title);
            pTitle.Format.HorizontalAlignment = HorizontalAlignment.Center;

            //-------------------- RECOMMENDATIONS --------------  
            string titleRecommendations = "RECOMENDATIONS:";
            generateParagraphList(section, titleRecommendations, audit, true);
            //----------------------- ISSUES --------------------
            string titleIssues = "ISSUES:";
            generateParagraphList(section, titleIssues, audit, false);

            // Save Word Document
            document.SaveToFile(@"D:\Credentials\WordDocumentGenerated.docx", FileFormat.Docx2013);

            Console.WriteLine("Documento Creado...!!!");
        }

        private static void generateParagraphList(Section section, string title, Audit audit, bool action)
        {
            Paragraph paragraph = section.AddParagraph();
            paragraph.AppendText(title);
            if (action)
            {
                foreach (var re in audit.Recommendations)
                {
                    // Add new paragraphs
                    Paragraph p = section.AddParagraph();
                    p.AppendText(re.Title + ": " + re.Description);
                    p.ListFormat.ApplyBulletStyle();
                    p.ListFormat.CurrentListLevel.NumberPosition = -10;
                }
            }
            else
            {
                foreach (var re in audit.Issues)
                {
                    // Add new paragraphs
                    Paragraph p = section.AddParagraph();
                    p.AppendText(re.Title + ": " + re.Description);
                    p.ListFormat.ApplyBulletStyle();
                    p.ListFormat.CurrentListLevel.NumberPosition = -10;
                }
            }
        }

        public static async Task<Audit> GetAudit()
        {
            Console.WriteLine("Loading information MockData API. . . . . . \n");

            Audit audit;
            using (var client = new HttpClient())
            {
                var response = await client.GetStringAsync("http://www.mocky.io/v2/5bbe0ac43100003800711390");
                audit = JsonConvert.DeserializeObject<Audit>(response);
            }
            return audit;
        }

        public static void CountParagraph(Document document)
        {
            Console.WriteLine(">>>>> Contar los Paragrphs <<<<<\n");
            Console.WriteLine("Este documento en la seccion 0 tiene: " + document.Sections[0].Paragraphs.Count + " Paragraph");
        }

        #region Interface
        public static void setConsoleTitle()
        {
            Console.WriteLine("                                         ===================================");
            Console.WriteLine("                                         AUDI REPORT WORD DOCUMENT GENERATOR");
            Console.WriteLine("                                         ===================================\n");
        }
        #endregion
    }
}
