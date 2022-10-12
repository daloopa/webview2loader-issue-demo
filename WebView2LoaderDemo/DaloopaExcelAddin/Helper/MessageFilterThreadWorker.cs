using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace DaloopaExcelAddin.Helper
{
    class MessageFilterThreadWorker : IMessageFilter
    {
        private Thread _thread;

        private IMessageFilter _oldMessageFilter;
        private MessageFilterThreadWorker _currentMessageFilter;

        public MessageFilterThreadWorker() { 
            _currentMessageFilter = this;
        }

        #region IMessageFilter Members
        [DllImport("ole32.dll")]
        static extern int CoRegisterMessageFilter(Helper.IMessageFilter lpMessageFilter, out Helper.IMessageFilter lplpMessageFilter);

        int IMessageFilter.HandleInComingCall(uint dwCallType, IntPtr htaskCaller, uint dwTickCount, INTERFACEINFO[] lpInterfaceInfo)
        {
            // We're the client, so we won't get HandleInComingCall calls.
            return 1;
        }

        int IMessageFilter.RetryRejectedCall(IntPtr htaskCallee, uint dwTickCount, uint dwRejectType)
        {
            // The client will get RetryRejectedCall calls when the main Excel
            // thread is blocked. We can handle this by attempting to retry
            // the operation. This will continue to fail so long as Excel is 
            // blocked.
            // As an alternative to simply retrying, we could put up
            // a dialog telling the user to close the other dialog (and the
            // new one) in order to continue - or to tell us if they want to
            // abandon this call
            // Expected return values:
            // -1: The call should be canceled. COM then returns RPC_E_CALL_REJECTED from the original method call.
            // Value >= 0 and <100: The call is to be retried immediately.
            // Value >= 100: COM will wait for this many milliseconds and then retry the call.
            String msg = String.Format("{0}@RetryRejectedCall", this.GetType().FullName);
            ThreadDiagnostics.PrintThreadInfo(msg);
            return 1;
        }

        int IMessageFilter.MessagePending(IntPtr htaskCallee, uint dwTickCount, uint dwPendingType)
        {
            return 1;
        }

        # endregion

        // It is important that we do not put any logic that modify UI in Action,
        // which will cause thread exceptions. Use this method to run CPU bound tasks
        // in background threads. Do not use this method to protect logics on the main
        // UI thread. See ShellViewModel for logic that protects UI level code.
        public void run(Action action) {
            _thread = new Thread(() =>
            {
                try
                {
                    int result = CoRegisterMessageFilter(_currentMessageFilter, out _oldMessageFilter);

                    action();
                }
                catch (COMException ex)
                {
                }
            });

            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();
            _thread.Join();
        }
    }
}
