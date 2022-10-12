using DaloopaExcelAddin.Helper;
using JetBrains.Annotations;
using MahApps.Metro.Controls.Dialogs;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;

namespace DaloopaExcelAddin.ViewModels
{
    public class BindableBase : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        protected virtual bool SetProperty<T>(ref T storage, T value, [CallerMemberName] string propertyName = "")
        {
            if (EqualityComparer<T>.Default.Equals(storage, value))
                return false;
            storage = value;
            this.OnPropertyChanged(propertyName);
            return true;
        }

        // This is how to use it
        //public string FirstName
        //{
        //    get { return _firstName; }
        //    set { SetProperty(ref _firstName, value); }
        //}
    }
    public class ShellViewModel : BindableBase,  IMessageFilter
    {
        #region Properties
        //public event EventHandler OnRequestClose;
        public CancellationTokenSource TokenSource { get; set; }
        public Action CloseAction { get; set; }
        private ObservableCollection<MenuItemViewModel> _menuItems;
        private ObservableCollection<MenuItemViewModel> _menuOptionItems;
        private IMessageFilter _previousMessageFilter;

        private int _selectedIndex;
        public int SelectedIndex
        {
            get => _selectedIndex;
            set => SetProperty(ref _selectedIndex, value);
        }


        private BrowserViewModel _loginVM;
        public BrowserViewModel LoginVM
        {
            get => _loginVM;
            set => SetProperty(ref _loginVM, value);
        }


        public IDialogCoordinator DialogCoordinator;
        #endregion

        #region Commands
        #endregion

        #region Ctor
        public ShellViewModel(int index)
        {
            TokenSource = new CancellationTokenSource();
            LoginVM = new BrowserViewModel();

            CreateMenuItems();
            SelectedIndex = index;


            int result = CoRegisterMessageFilter(this, out _previousMessageFilter);

            ThreadDiagnostics.PrintThreadInfo("CoRegisterMessageFilter result: " + result);
        }
        #endregion

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

            // ShellViewModel implements IMessageFilter to handles "Application is busy error" on the main UI thread.
            // Another place where IMessageFilter is implmented is MessageFilterThreadWorker which is used to run background tasks.

            String msg = String.Format("{0}@RetryRejectedCall", this.GetType().FullName);
            ThreadDiagnostics.PrintThreadInfo(msg);
            return 1;
        }

        int IMessageFilter.MessagePending(IntPtr htaskCallee, uint dwTickCount, uint dwPendingType)
        {
            return 1;
        }
        #endregion


        #region Methods



        #endregion

        #region Menu Options
        public void CreateMenuItems()
        {
            MenuItems = new ObservableCollection<MenuItemViewModel>
            {
                LoginVM,
            };

            MenuOptionItems = new ObservableCollection<MenuItemViewModel>
            {
            };
        }

        public ObservableCollection<MenuItemViewModel> MenuItems
        {
            get => _menuItems;
            set => SetProperty(ref _menuItems, value);
        }

        public ObservableCollection<MenuItemViewModel> MenuOptionItems
        {
            get => _menuOptionItems;
            set => SetProperty(ref _menuOptionItems, value);
        }

        #endregion
    }
}
