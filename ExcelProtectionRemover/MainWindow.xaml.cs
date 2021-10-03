using System.Threading.Tasks;
using System.Windows;
using ExcelProtectionRemover.Process;
using Microsoft.Win32;

namespace ExcelProtectionRemover
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
        }

        private void ChooseFileButton_Click(object sender, RoutedEventArgs e)
        {
            FileDialog file = new OpenFileDialog();
            var result = file.ShowDialog();
            if (result.HasValue && !result.Value) return;
            FileNameTbx.Text = file.FileName;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var fileFullPath = FileNameTbx.Text;
            Task.Run(() =>
            {
                 var processor = new FileProcessor();
                 var result = processor.Process(fileFullPath);
            }).ContinueWith(x =>
            {
                MessageBox.Show("The file is converted!");
            });
        }
    }
}
