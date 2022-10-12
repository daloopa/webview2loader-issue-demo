using ExcelDna.Integration;
using System;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace DaloopaExcelAddin
{
    public class AddinModule : IExcelAddIn
    {

        public void AutoOpen()
        {
            System.Windows.Forms.Application.ThreadException +=
              new System.Threading.ThreadExceptionEventHandler(Application_ThreadException);

            // Add handler to handle the exception raised by additional threads
            AppDomain.CurrentDomain.UnhandledException +=
            new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);

            var comAddIn = new DaloopaComAddIn();
            ExcelComAddInHelper.LoadComAddIn(comAddIn);

            var version = Assembly.GetExecutingAssembly().GetName().Version;

            Application.EnableVisualStyles();

            SetCurrentInstance(this);

        }

        static void Application_ThreadException
                (object sender, System.Threading.ThreadExceptionEventArgs e)
        {// All exceptions thrown by the main thread are handled over this method

            ShowExceptionDetails(e.Exception);
        }

        static void CurrentDomain_UnhandledException
                (object sender, UnhandledExceptionEventArgs e)
        {// All exceptions thrown by additional threads are handled in this method

            ShowExceptionDetails(e.ExceptionObject as Exception);

            // Suspend the current thread for now to stop the exception from throwing.
            Thread.CurrentThread.Suspend();
        }

        static void ShowExceptionDetails(Exception Ex)
        {
            // Do logging of exception details
            MessageBox.Show(Ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public void AutoClose()
        {
        }

        private static AddinModule _currentInstance;

        public static AddinModule GetCurrentInstance()
        {
            return _currentInstance;
        }

        public static void SetCurrentInstance(AddinModule value)
        {
            _currentInstance = value;
        }

        public string EncodeTo64String(string plainText)
        {
            var plainTextBytes = Encoding.UTF8.GetBytes(plainText);
            return Convert.ToBase64String(plainTextBytes);
        }

        public string DecodeFrom64String(string base64EncodedData)
        {
            if (string.IsNullOrEmpty(base64EncodedData))
                return base64EncodedData;
            var base64EncodedBytes = Convert.FromBase64String(base64EncodedData);
            return Encoding.UTF8.GetString(base64EncodedBytes);
        }

        public Excel._Application ExcelApp
        {
            get
            {
                return (ExcelDnaUtil.Application as Excel._Application);
            }
        }

        public Excel._Application HostApplication { get; private set; }

        static NativeWindow xlNW = new NativeWindow();

        public System.Threading.AutoResetEvent LoginThreadDone = new System.Threading.AutoResetEvent(false);

    }
}
