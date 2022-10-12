using DaloopaExcelAddin.Views;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Web.WebView2.Core;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Interop;
using System.Windows.Threading;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace DaloopaExcelAddin
{
    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            var loaderFolderUri = $"{AppDomain.CurrentDomain.BaseDirectory}/runtimes/win-{RuntimeInformation.ProcessArchitecture}/native";

            CoreWebView2Environment.SetLoaderDllFolderPath(loaderFolderUri);
            return @"
      <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
      <ribbon>
        <tabs>
          <tab idMso='TabHome'>
            <group id='group1' label='Daloopa' insertAfterMso='GroupFont'>
                <button id='btn_browse' label='Browse' imageMso='ArrangeByAccount' size='large' onAction='HandleBrowseButtonClick'/>
            </group>
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
        }

        private ShellView form;
        public void HandleBrowseButtonClick(IRibbonControl control)
        {
            ShowPopup(0);
        }


        public bool ShowPopup(int index)
        {
            if (form != null && form.IsVisible) return false;

            var xlApp = (Application)ExcelDnaUtil.Application;

            var hwind = xlApp.Hwnd;

            var thread = new Thread(() =>
            {

                form = new ShellView(index);
                WindowInteropHelper windowInteropHelper = new WindowInteropHelper(form)
                {
                    Owner = new IntPtr(hwind)
                };

                form.Show();

                Dispatcher.Run();
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            return false;
        }

    }
}
