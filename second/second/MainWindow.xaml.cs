using System;
using System.Collections.Generic;
using System.Diagnostics;
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
using AutoIt;
using System.IO;
using System.Runtime.InteropServices;

namespace second
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern bool PostMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        [DllImport("user32.dll")]
        static extern int LoadKeyboardLayout(string pwszKLID, uint Flags);

        public MainWindow()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            string lang = "00000419";
            int ret = LoadKeyboardLayout(lang, 1);
            PostMessage(GetForegroundWindow(), 0x50, 1, ret);
            var exePath = AppDomain.CurrentDomain.BaseDirectory;
            object path = Path.Combine(exePath, "parse_status.exe");
            string url = text_url.Text;
            if (url.IndexOf("m.vk") == -1)
            {
                int i = url.IndexOf("vk.com");
                url = url.Insert(i, "m.");
            }
            AutoItX.ClipPut(url);
            Process.Start("IExplore.exe");
            AutoItX.AutoItSetOption("WinTitleMatchMode", 2);
            AutoItX.WinWaitActive("Internet Explorer");
            AutoItX.Send(url + "{ENTER}");
            Process.Start(@"" + path);
            AutoItX.WinWaitActive("MainWindow");
            lb_status.Content = AutoItX.ClipGet();
        }
    }
}
