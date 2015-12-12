using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WordToPDF
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void WordBtn_Click(object sender, RoutedEventArgs e)
        {
            
            //WordConvertPDF.WordToPDFHelper.WordToPDF(@"C:\Users\DNT\Desktop\2.docx", @"C:\Users\DNT\Desktop\1.pdf");
            string log;
            WordConvertPDF.WordToPDFHelper.WordToPDFCreateLog(@"C:\Users\DNT\Desktop\1.docx", @"C:\Users\DNT\Desktop\3.pdf", out log);
            MessageBox.Show(log);
        }

        private void PdfBtn_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
