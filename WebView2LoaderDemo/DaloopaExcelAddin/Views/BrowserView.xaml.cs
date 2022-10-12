using Microsoft.Web.WebView2.Core;
using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace DaloopaExcelAddin.Views
{
    /// <summary>
    /// Interaction logic for BrowseView.xaml
    /// </summary>
    public partial class BrowseView : UserControl
    {
        public BrowseView()
        {
            InitializeComponent();
        }
        bool isloaded = false;
        public async void InitializeWebView2Async()
        {
            try
            {
                // must create a data folder if running out of a secured folder that can't write like Program Files
                var env = await CoreWebView2Environment.CreateAsync(userDataFolder: Path.Combine(Path.GetTempPath(), "Daloopa_Browser"));
                await webView.EnsureCoreWebView2Async(env);
                webView.CoreWebView2.Navigate("https://www.google.com/");
            }
            catch (Exception e)
            {
                MessageBox.Show("Error " + e.Message);
            }
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            if (!isloaded)
            {
                InitializeWebView2Async();
                isloaded = true;
            }

        }
    }

}
