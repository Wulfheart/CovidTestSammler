using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using MsgReader.Outlook;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection.Metadata.Ecma335;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CovidTestOutlookCollector
{
    class Entity
    {
        public string Surname { get; set; }
        public string Forename { get; set; }
        public string Gender { get; set; }
        public DateTime Sampling { get; set; }
        public string Barcode { get; set; }
        public string Result { get; set; }
        public string PdfPath { get; set; }
        public string Password { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string Birthdate { get; set; }

        // This could also go wrong sometime due to too much memory allocation
        private Storage.Attachment attachment { get; set; }

        public Entity(PasswordMsg pwmsg, PdfMsg pdfmsg)
        {
            Password = pwmsg.Password;
            attachment = pdfmsg.Message.Attachments.Where(a => a.GetType().Equals(typeof(MsgReader.Outlook.Storage.Attachment))).Cast<MsgReader.Outlook.Storage.Attachment>().Where(a => a.FileName.Contains(pwmsg.PdfFileName, StringComparison.OrdinalIgnoreCase)).First();

            Barcode = pwmsg.ID;

        }
        public void ParseData()
        {
            MemoryStream ms = new MemoryStream(attachment.Data);
            PdfDocument pdf = new PdfDocument(new PdfReader(ms, new ReaderProperties().SetPassword(System.Text.ASCIIEncoding.ASCII.GetBytes(Password))));
            string text = PdfTextExtractor.GetTextFromPage(pdf.GetFirstPage());
            Surname = new Regex(@"(?<=\bLastname\s)(\w+)").Match(text).Value;
            Forename = new Regex(@"(?<=\bFirstname\s)(\w+)").Match(text).Value;
            Gender = new Regex(@"(?<=\bSex\s)(\w+)").Match(text).Value;
            Birthdate = new Regex(@"(?<=\bDate of Birth\s)(.+)").Match(text).Value;
            string tmpResult = new Regex(@"(?<=\bResult:\s)(\w+)").Match(text).Value;
            string tmpBarcode = new Regex(@"(?<=\bBarcode\/ Barcode\s)(\w+)").Match(text).Value;
            string tmpSampling = new Regex(@"(?<=\bSampling\s)(.+)").Match(text).Value.Replace("(CEST)", "").Replace("(CET)", "");
            if (tmpBarcode != Barcode)
            {
                pdf.Close();
                throw new Exception("Es ist eine Ungereimtheit aufgetreten zwischen dem Barcode, der aus dem Betreff entnommen wurde, und der, der aus der PDF ausgelesen wurde.");
            }
            Sampling = DateTime.Parse(tmpSampling);
            if (String.Equals(tmpResult, "negativ", StringComparison.OrdinalIgnoreCase))
            {
                Result = "negativ";
            }
            else if (String.Equals(tmpResult, "positiv", StringComparison.OrdinalIgnoreCase))
            {
                Result = "positiv";
            }
            else
            {
                Result = "nicht eindeutig";
            }
            Address = ReadLine(text, 5);
            City = ReadLine(text, 6);
            pdf.Close();


        }

        public void Save(string folder)
        {
            PdfPath = Path.Combine(folder, $"{Surname}_{Forename}_Ergebnis_von_{Sampling.ToString("yyyy-MM-dd")}.pdf");
            FileInfo file = new FileInfo(PdfPath);
            file.Directory.Create();
            MemoryStream ms = new MemoryStream(attachment.Data);
            PdfDocument pdf = new PdfDocument(new PdfReader(ms, new ReaderProperties().SetPassword(System.Text.ASCIIEncoding.ASCII.GetBytes(Password))), new PdfWriter(file));
            pdf.Close();
        }

        public void WriteToExcel(string folder, string signature)
        {
            // Create Excel if not exists
            FileInfo fi = new FileInfo(Path.Combine(folder, signature));
            ExcelPackage p = new ExcelPackage(fi);

            ExcelWorksheet ws = p.Workbook.Worksheets["Daten"];
            if (ws == null)
            {
                ws = p.Workbook.Worksheets.Add("Daten");
                ws.Cells[1, 1].Value = "Nachname";
                ws.Cells[1, 2].Value = "Vorname";
                ws.Cells[1, 3].Value = "Geschlecht";
                ws.Cells[1, 4].Value = "Testdatum";
                ws.Cells[1, 5].Value = "Barcode";
                ws.Cells[1, 6].Value = "Ergebnis";
                ws.Cells[1, 7].Value = "Geburtsdatum";
                ws.Cells[1, 8].Value = "Adresse";
                ws.Cells[1, 9].Value = "Stadt";
                ws.Cells[1, 10].Value = "Pdf";
            }
            ws.Cells.AutoFitColumns(15, 50);

            int row = ws.Dimension.Rows + 1;
            ws.Cells[row, 1].Value = Surname;
            ws.Cells[row, 2].Value = Forename;
            ws.Cells[row, 3].Value = Gender;
            ws.Cells[row, 4].Value = Sampling.ToString("dd.MM.yyyy");
            ws.Cells[row, 5].Value = Barcode;
            ws.Cells[row, 6].Value = Result;
            ws.Cells[row, 7].Value = Birthdate;
            ws.Cells[row, 8].Value = Address;
            ws.Cells[row, 9].Value = City;
            ws.Cells[row, 10].Value = PdfPath;

            p.Save();
        }

        private string ReadLine(string text, int lineNumber)
        {
            var reader = new StringReader(text);

            string line;
            int currentLineNumber = 0;

            do
            {
                currentLineNumber += 1;
                line = reader.ReadLine();
            }
            while (line != null && currentLineNumber < lineNumber);

            return (currentLineNumber == lineNumber) ? line :
                                                       string.Empty;
        }



    }
}
