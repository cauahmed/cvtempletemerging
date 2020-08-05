using System;
using System.Collections.Generic;
using System.Text;

using Aspose.Words;
using Aspose.Words.Saving;

namespace TestApplication
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document("C:\\Personal Development\\Resume_new-20200317T114937Z-001\\Resume_new\\SALES\\JuniorSalesCadet.docx");
            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.Compliance = PdfCompliance.PdfA1b;
            doc.Save("C:\\Personal Development\\Resume_new-20200317T114937Z-001\\Resume_new\\SALES\\JuniorSalesCadet.pdf", saveOptions);
	
            //Adding array of objects to merge fields using aspose
	    /*const string dataDir = "C:\\Personal Development\\CV merging\\Data\\";
            Document doc = new Document(dataDir + "CV Template.docx");
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;
            doc.MailMerge.Execute(new string[] { "FullName", "Company", "Address", "Address2", "City", "MyImage"}, new object[] {"Asif Ahmed", "UTAS", "Carlingford", "", "Australia", File.ReadAllBytes(dataDir + "usericon.png")});
            doc.Save(dataDir + "CV Out.docx");*/

        }
    }
}
