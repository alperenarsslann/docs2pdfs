using System;
using System.Windows.Forms;

namespace DocxToPdf
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                MessageBox.Show("Dosya belirtilmedi.", "DocxToPdf");
                return;
            }

            string docxPath = args[0];

            if (!System.IO.File.Exists(docxPath))
            {
                MessageBox.Show($"Dosya bulunamadi:\n{docxPath}", "DocxToPdf");
                return;
            }

            try
            {
                if (Converter.ConvertToPdf(docxPath))
                {
                    string pdfPath = System.IO.Path.ChangeExtension(docxPath, ".pdf");
                    MessageBox.Show($"PDF olusturuldu:\n{pdfPath}", "DocxToPdf");
                }
                else
                {
                    MessageBox.Show("Donusum basarisiz.", "DocxToPdf");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Hata:\n{ex.Message}", "DocxToPdf");
            }
        }
    }
}