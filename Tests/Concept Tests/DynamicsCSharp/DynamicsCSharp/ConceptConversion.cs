//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Windows.Forms;
//using NetOffice;
//using Excel = NetOffice.ExcelApi;

//namespace DynamicsCSharp
//{
//    internal class ConceptConversion
//    {
//        internal void Test()
//        {
//            try
//            {               
//                Type excelType = System.Type.GetTypeFromProgID("Excel.Application", true);
//                object interopProxy = Activator.CreateInstance(excelType);

//                COMDynamicObject.TryConvertFailResult = true;
//                dynamic application = new COMDynamicObject(interopProxy);

//                application.Visible = true;
//                application.Workbooks.Add();
//                application.Workbooks.Add();
//                application.Workbooks.Add();

//                Console.WriteLine("Proxy Count(must 7) {0}", NetOffice.Core.Default.ProxyCount);

//                //// compare and indexers
//                bool book1Active = application.ActiveWorkbook == application.Workbooks[1];
//                bool book3Active = application.ActiveWorkbook == application.Workbooks[3];
//                Console.WriteLine("book1Active {0} book3Active {1}", book1Active, book3Active);
//                Console.WriteLine("Proxy Count(must 13) {0}", NetOffice.Core.Default.ProxyCount);

//                foreach (var book in application.Workbooks)
//                {
//                    Console.WriteLine(book.Name);
//                    foreach (var sheet in book.Sheets)
//                    {
//                        Console.WriteLine(sheet.Name);
//                    }
//                }

//                // implicit conversions => COMDynamicObject::DynamicObject::TryConvert
//                // Confusing stuff about dynamic and implicit/explicit conversions to read here:
//                // https://stackoverflow.com/questions/3492955/dynamicobject-tryconvert-not-called-when-casting-to-interface-type
//                // Not sure what means John Skeet here to handle that better
//                // with IDynamicMetaObjectProvider to bring explicit conversion to work

//                Console.WriteLine("Convert dynamic back to strong type.");
//                Console.WriteLine("This will create 2 new instances in NetOffice com proxy management");
//                COMObject comApplication = application;
//                Excel.Application excelApplication = application;
//                Console.WriteLine("Proxy Count(must 35) {0}", NetOffice.Core.Default.ProxyCount);

//                Excel.Workbook failedWorkBook = application; // must fail because application isn't a workbook
//                ICOMObject wantedRootInterface = application; // no TryConvert required
//                COMObject wantedCOMObject = application.ActiveWorkbook;
//                string wantedString = application;
//                object objectApplication = application; // no TryConvert required
//                IDisposable wantedDisposable = application; // no TryConvert required

//                // Explicit conversion doesnt work - any idea ???
//                Excel.Application failedApplication = application as Excel.Application;

//                application.Quit();

//                // the converted applications are also root proxies
//                if (null != comApplication)
//                    comApplication.Dispose();
//                if (null != excelApplication)
//                    excelApplication.Dispose();

//                application.Dispose();
//            }
//            catch (Exception exception)
//            {
//                Console.WriteLine(exception);
//            }

//            Console.WriteLine("Press any key.");
//            Console.ReadKey();
//        }
//    }
//}
