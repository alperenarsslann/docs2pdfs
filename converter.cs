using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;

namespace DocxToPdf
{
    public class Converter
    {
        public static bool ConvertToPdf(string docxPath)
        {
            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                string pdfPath = System.IO.Path.ChangeExtension(docxPath, ".pdf");

                wordApp = new Word.Application();
                wordApp.Visible = false;
                wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

                doc = wordApp.Documents.Open(docxPath);
                doc.SaveAs2(pdfPath, Word.WdSaveFormat.wdFormatPDF);

                return true;
            }
            catch (System.Exception ex)
            {
                System.Console.WriteLine($"Hata: {ex.Message}");
                return false;
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close(false);
                    Marshal.ReleaseComObject(doc);
                }
                if (wordApp != null)
                {
                    wordApp.Quit(false);
                    Marshal.ReleaseComObject(wordApp);
                }
            }
        }
    }
}