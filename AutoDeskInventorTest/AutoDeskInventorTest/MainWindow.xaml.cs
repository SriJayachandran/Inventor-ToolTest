using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace AutoDeskInventorTest
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void selectsrctemplatebtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            DialogResult result = dialog.ShowDialog();
            srcfldrtxtbox.Text = dialog.FileName.ToString();
            String fldrpathlen = srcfldrtxtbox.Text;
            if (fldrpathlen.Length == 0)
                System.Windows.MessageBox.Show(this, "Please select source template file");
        }

        private void selectsrcExcelbtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            DialogResult result = dialog.ShowDialog();
            srcExceltxtbox.Text = dialog.FileName.ToString();
            String excelFile = srcExceltxtbox.Text;
            if (excelFile.Length == 0)
                System.Windows.MessageBox.Show(this, "Please select input excel file");
        }

        private void btnCreateDrawing_Click(object sender, RoutedEventArgs e)
        {
            coreFunctionality.toolTest(srcExceltxtbox.Text, srcfldrtxtbox.Text);
        }
    }
}
