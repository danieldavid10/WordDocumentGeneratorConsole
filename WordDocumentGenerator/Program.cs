using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
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
            document.LoadFromFile(@"D:\WDocuments\WordDocumentGenerated.docx");

            //Generar nuevo documento word
            //GenerateDocument(BlankDocument);

            // contar cuantos parrafos tiene una seccion
            //CountParagraph(document);

            // Ocultar texto
            //HidenText(document);

            LoadDataInJsonFile(document);


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
            document.SaveToFile(@"D:\WDocuments\WordDocumentGenerated.docx", FileFormat.Docx2013);

            // Open Document
            Process.Start(@"D:\WDocuments\WordDocumentGenerated.docx");

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

        public static void HidenText(Document doc)
        {
            Section sec = doc.Sections[0];
            Paragraph para = sec.Paragraphs[0];

            (para.ChildObjects[0] as TextRange).CharacterFormat.Hidden = false;
            // save
            doc.SaveToFile(@"D:\WDocuments\WordDocumentGenerated.docx", FileFormat.Docx2013);
        }

        public static void LoadDataInJsonFile(Document document)
        {
            Console.WriteLine(">>>>> Cargar Datos en el JSON <<<<<\n");

            Audit audit = new Audit();
            List<Recommendation> recommendations = new List<Recommendation>();
            List<Issue> issues = new List<Issue>();

            var paragraphs = document.Sections[0].Paragraphs;
            int i = 1;

            audit.AuditId = 1;
            audit.Title = paragraphs[0].Text;

            if (paragraphs[i].Text == "RECOMENDATIONS:")
            {
                i++;
                while (paragraphs[i].Text != "ISSUES:")
                {
                    Recommendation recom = new Recommendation();
                    recom.RecommendationId = 1;
                    recom.Title = "Title of Recomendation";
                    recom.Description = paragraphs[i].Text;
                    i++;

                    recommendations.Add(recom);
                }
                i++;
                while (i < paragraphs.Count)
                {
                    Issue issue = new Issue();
                    issue.IssueId = 1;
                    issue.Title = "Title of Issue";
                    issue.Description = paragraphs[i].Text;
                    i++;

                    issues.Add(issue);
                }
            }

            audit.Recommendations = recommendations;
            audit.Issues = issues;

            JsonContentGenerator(audit);
        }

        private static void JsonContentGenerator(Audit audit)
        {
            Console.WriteLine(">>>>> Json Content <<<<<\n");
            Console.WriteLine("{\n");
            Console.WriteLine("  \nId: " + audit.AuditId);
            Console.WriteLine("  \nTitle: " + audit.Title);
            Console.WriteLine("  \nRecomendations: \n     [\n");

            foreach (var re in audit.Recommendations)
            {
                Console.WriteLine("{\n");
                Console.WriteLine("  \nRecomendationId: " + re.Title);
                Console.WriteLine("  \nTitle: " + re.Title);
                Console.WriteLine("  \nDescription: " + re.Description);
                Console.WriteLine("\n{\n");
            }
            foreach (var iss in audit.Issues)
            {
                Console.WriteLine("{\n");
                Console.WriteLine("  \nRecomendationId: " + iss.Title);
                Console.WriteLine("  \nTitle: " + iss.Title);
                Console.WriteLine("  \nDescription: " + iss.Description);
                Console.WriteLine("\n{\n");
            }
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
