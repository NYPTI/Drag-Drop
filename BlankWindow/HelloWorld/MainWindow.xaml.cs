using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Media;

namespace HelloWorld
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

        private void btnClick(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/NYPTI/Outlook-Drag-Drop/releases");
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            VerifyProgress.Foreground = Brushes.Green;
            OutputTextBox.Text = "Verification in process...\n";
            VerifyProgress.Value = 1;
            string s86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86) + "\\New York Prosecutors Training Institute\\NYPTI Outlook Drag Drop (x86)\\";
            string s64 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + "\\New York Prosecutors Training Institute\\NYPTI Outlook Drag Drop (x64)\\";
            bool i86 = false;
            bool i64 = false;
            VerifyProgress.Value = 10;
            try
            {
                IEnumerable<string> files = Directory.EnumerateFiles(s86);
                OutputTextBox.Text += "Found 32 bit files:\n";
                foreach (string path in files)
                {
                    OutputTextBox.Text += path + "\n";
                }
                i86 = true;
            }
            catch { }
            VerifyProgress.Value = 50;
            try
            {
                IEnumerable<string> files = Directory.EnumerateFiles(s64);
                OutputTextBox.Text += "Found 64 bit files:\n";
                foreach (string path in files)
                {
                    OutputTextBox.Text += path + "\n";
                }
                i64 = true;
            }
            catch
            { }
            VerifyProgress.Value = 90;
            if (i86)
            {
                OutputTextBox.Text += "\nThe 32 bit version is successfully installed.\n";
            }
            if (i64)
            {
                OutputTextBox.Text += "\nThe 64 bit version is successfully installed.\n";
            }
            if (!i86 && !i64)
            {
                OutputTextBox.Text += "Could not detect any installation of NYPTI Outlook Drag Drop.";
                VerifyProgress.Foreground = Brushes.Red;
            }
            VerifyProgress.Value = 100;
        }
    }
}
