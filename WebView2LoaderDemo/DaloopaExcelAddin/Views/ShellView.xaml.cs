using DaloopaExcelAddin.ViewModels;
using MahApps.Metro.Controls;

namespace DaloopaExcelAddin.Views
{
    /// <summary>
    /// Interaction logic for ShellView.xaml
    /// </summary>
    public partial class ShellView : MetroWindow
    {
        public ShellViewModel shellViewModel { get; set; }
        public ShellView(int index)
        {
            InitializeComponent();
            shellViewModel = new ShellViewModel(index);
            DataContext = shellViewModel;
        }

    }
}
