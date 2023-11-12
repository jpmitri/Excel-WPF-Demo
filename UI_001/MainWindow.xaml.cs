using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Forms;

using Excel_Functions;

using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace UI_001
{
    /// <summary>
    ///     Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly List<string> Object_Type         = new() {"Inbound", "Outbound", "Special Projects"};
        private          string       Destination         = "";
        private          List<string> Selected_Files_List = new();
        public MainWindow()
        {
            InitializeComponent();
            Option_Selector.ItemsSource   = Object_Type;
            Option_Selector.SelectedIndex = 0;
        }
        private void select_File(object          sender,
                                 RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new()
            {
                Title            = "Select The Required Excel Files",
                Filter           = "Excel files (*.xlsx)|*.xlsx|Excel files (*.xlx)|*.xlx|All files (*.*)|*.*",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                Multiselect      = true
            };
#if DEBUG
            openFileDialog.InitialDirectory =
                "F:\\001 - GEMICo\\002 - Mado\\Project 1\\003 - Documents\\Inbound-20230630T193845Z-001\\Inbound";
#endif
            if (openFileDialog.ShowDialog() == true)
            {
                Selected_Files.Text = string.Join(";\n", openFileDialog.FileNames);
                Selected_Files_List = openFileDialog.FileNames.ToList<string>();
            }
        }
        private void save_File(object          sender,
                               RoutedEventArgs e)
        {
            FolderBrowserDialog saveFileDialog       = new();
            DialogResult        saveFileDialogResult = saveFileDialog.ShowDialog();
            if (saveFileDialogResult.ToString().ToLower().Trim().Equals("ok"))
            {
                Destination = saveFileDialog.SelectedPath;
                Excel ex = new(Selected_Files_List);
                _ = ex.writeInbound(Destination, Joker_Value.Text);
            }
        }
        private enum Options { Inbound, Outbound, Special_Projects }
    }
}
