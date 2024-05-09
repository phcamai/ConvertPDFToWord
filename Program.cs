// See https://aka.ms/new-console-template for more information

using Acrobat;
using System.Globalization;
using System.Reflection;

public interface IControl
{
    void Paint();
}

public class SampleClass : IControl
{
    // Paint() is inherited from IControl.
    public void Paint()
    {

        String szPdfPathConst = "D:\\1.pdf";  //Application.StartupPath + "\\..\\..\\Data\\Department\\01_CoverSheet_CityFoneDirectory.pdf";
        String szDocxPathConst = "D:\\1.docx";

        Console.WriteLine("Open document: %s", szPdfPathConst);

        //Initialize Acrobat by creating App object
        AcroAppClass mApp = new AcroAppClass();


        //Show Acrobat
        //mApp.Show();
        mApp.Hide();

        //set AVDoc object
        AcroAVDocClass avDoc = new AcroAVDocClass();

        //open the PDF
        if (avDoc.Open(szPdfPathConst, ""))
        {
            CAcroPDDoc pdfd = (CAcroPDDoc)avDoc.GetPDDoc();
            Object jsObj = pdfd.GetJSObject();
            Type jsType = pdfd.GetType();
            //have to use acrobat javascript api because, acrobat
            object[] saveAsParam = { szDocxPathConst, "com.adobe.acrobat.doc", "", false, false };
            jsType.InvokeMember("saveAs", /*BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance*/BindingFlags.Default, null, jsObj, saveAsParam, CultureInfo.InvariantCulture);
        }

        Console.WriteLine("Convert pdf to docx: %s", szDocxPathConst);

        // close main doc
        avDoc.Close(1);  // '1' parameter means close without saving...a '0' would prompt the user to save

        if (mApp != null)
        {
            mApp.CloseAllDocs();
            mApp.Exit();
        }

    }
}

SampleClass sampleClass = new SampleClass();


