using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.Extensibility;
using ExcelDna.Integration.CustomUI;
using Excel = Microsoft.Office.Interop.Excel;

namespace DaloopaExcelAddin
{
    [ComVisible(true)]
    public class DaloopaComAddIn : ExcelComAddIn
    {
        private Excel.Application _applicationObject;
        private object _addInInstance;
        private Process _mXlProcess;

        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        public DaloopaComAddIn()
        {
        }

        public override void OnConnection(object application, 
            ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _applicationObject = (Excel.Application)application;
            _addInInstance = addInInst;
            _mXlProcess = GetExcelProcess(_applicationObject);
        }

        public override void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            //SWF.MessageBox.Show("OnDisconnection");
        }

        public override void OnAddInsUpdate(ref Array custom)
        {
            //SWF.MessageBox.Show("OnAddInsUpdate");
        }

        public override void OnStartupComplete(ref Array custom)
        {
            //SWF.MessageBox.Show("OnStartupComplete");
        }

        public override void OnBeginShutdown(ref Array custom)
        {
            if(_applicationObject != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_applicationObject);

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();

            if (!_mXlProcess.WaitForExit(5000))
            {
                // This should never happen, Unexpected zombie Excel
                _mXlProcess.Kill();
            }
        }

        Process GetExcelProcess(Excel.Application excelApp)
        {
            int id;
            GetWindowThreadProcessId(excelApp.Hwnd, out id);
            return Process.GetProcessById(id);
        }
    }
}
