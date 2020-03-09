using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Pdf;

namespace etaxOneth_Printer
{
    public class Class1
    {
        public void printer()
        {
            PrintMethod("D:\\Work Arena\\test.pdf", "Microsoft Print to PDF", 2);

        }
        public void PrintMethod(string path,string printer_name,short copies)
        {
            Console.WriteLine(path + " " + printer_name + " " + copies + " _printer");
            PrinterSettings oPrinterSettings = new PrinterSettings();
            PdfDocument pdfdocument = new PdfDocument();
            try
            {                
                pdfdocument.LoadFromFile(path);
                pdfdocument.PrinterName = printer_name;
                pdfdocument.PrintDocument.PrinterSettings.Copies = copies;
                pdfdocument.PrintDocument.Print();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                pdfdocument.Dispose();
            }            
            
        }
    }
}
