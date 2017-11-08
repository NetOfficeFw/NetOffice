using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NetOffice;

namespace DynamicsCSharp
{
    internal class ConceptOutArgs
    {
        internal void Test()
        {
            try
            {
                Type wordType = System.Type.GetTypeFromProgID("Word.Application", true);
                object interopProxy = Activator.CreateInstance(wordType);

                COMDynamicObject.TryConvertFailResult = true;
                dynamic application = new COMDynamicObject(interopProxy);
                application.Visible = true;

                application.DisplayAlerts = 0;
                var document = application.Documents.Add();
                application.Selection.TypeText("Hello World");

                int left = 0;
                int top = 0;
                int width = 0;
                int height = 0;

                dynamic window = application.ActiveWindow;
                dynamic range = application.Selection.Range;
                window.GetPoint(out left, out top, out width, out height, range);
                MessageBox.Show(string.Format("GetPoint returns Left:{0} Top:{1} Width:{2} Height:{3}", left, top, width, height));

                document.Saved = true;
                application.Quit();
                application.Dispose();
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
            }

            Console.WriteLine("Press any key.");
            Console.ReadKey();
        }
    }
}
