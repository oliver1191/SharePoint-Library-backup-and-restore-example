using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _2018081301MyLibraryDocument
{
    class Program
    {
        static void Main(string[] args)
        {
            string myServiceUrl = @"http://dragonma:45847";
            string myClientUrl = @"http://dragonma:19004";
            //string myXmlPath = @"C:\Users\administrator.SPCARTOON\Desktop\XML\2018081301";
            string destinationPath = @"C:\Users\administrator.SPCARTOON\Desktop\XML";
            string myDestinationDirectoryName = @"myLibrary Document";
            string myFileName = "2018081301Test";
            string xmlPath = "C:\\Users\\administrator.SPCARTOON\\Desktop\\XML\\MySharePointXML2";
            MyLibraryDocument myLibraryDocument = new MyLibraryDocument();
            //myLibraryDocument.WriteXml(myServiceUrl, @"Shared Documents");            
            //myLibraryDocument.ReadXml(myClientUrl, @"Shared Documents", xmlPath);
            //myLibraryDocument.CreateDirectoryTree(myServiceUrl, destinationPath);
            //myLibraryDocument.WriteLog("test");    
            myLibraryDocument.showListColumn(myClientUrl, "MyList201809050102");      
            Console.WriteLine("ssss");
            Console.ReadKey();
        }
    }
}
