using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace DaloopaExcelAddin.Helper
{
    class ThreadDiagnostics
    {
        private static Object lockObj = new Object();

        public static void PrintThreadInfo(string taskName)
        {
            Thread _thread = Thread.CurrentThread;

            String msg = null;

            lock (lockObj)
            {
                msg = String.Format("######### ThreadDiagnostics ##########\n") +
                    String.Format("{0} thread information\n", taskName) +
                   String.Format(" Background: {0}\n", _thread.IsBackground) +
                   String.Format(" Thread Pool: {0}\n", _thread.IsThreadPoolThread) +
                   String.Format(" Thread ID: {0}\n", _thread.ManagedThreadId) +
                   String.Format(" Thread ApartmentState: {0}\n", _thread.ApartmentState);
            }

            Debug.WriteLine(msg);
            Debug.WriteLine(Environment.StackTrace);
        }
    }
}
