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
            //打开对话框
            Microsoft.Win32.OpenFileDialog dialog = new Microsoft.Win32.OpenFileDialog();
            //过滤器
            dialog.Filter = "文档类型|*.doc;*.docx";
            //允许选多个文件
            dialog.Multiselect = true;
            //是否打开
            var result = dialog.ShowDialog();
            if (result == true)
            {
                //总计
                int total = 0;
                //保存路径
                string savePath = string.Empty;
                //获得文件
                string[] files = dialog.FileNames;
                for (int i = 0; i < files.Length; i++)
                {
                    savePath = System.IO.Path.GetDirectoryName(files[0]);
                    //文件名
                    string saveName = System.IO.Path.GetFileNameWithoutExtension(files[0]) + "pdf";
                    //存在就转换
                    if (System.IO.File.Exists(files[i]))
                    {
                        if (WordConvertPDF.WordToPDFHelper.WordToPDF(files[0], savePath + saveName))
                        {
                            total++;
                        }
                    }
                }
                MessageBox.Show("成功转换了：" + total + " 个文件", "逆天友情提醒");
                if (!string.IsNullOrEmpty(savePath))
                {
                    System.Diagnostics.Process.Start(savePath);
                }
            }
        }

        private void PdfBtn_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("V1.1版本的时候会公开~敬请期待","逆天友情提醒");
            System.Diagnostics.Process.Start("http://www.dkill.net");
        }

        private void TextBlock_MouseDown(object sender, MouseButtonEventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/dunitian/WordConvertPDF");
        }
    }
}
