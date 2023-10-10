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
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private List<String> Selected_Files_List = new();
        private readonly List<String> Object_Type = new() { "Inbound", "Outbound", "Special Projects" };
        private String Destination = "";

        private enum Options
        {
            Inbound,
            Outbound,
            Special_Projects
        }
        public MainWindow()
        {
            InitializeComponent();
            Option_Selector.ItemsSource = Object_Type;
            Option_Selector.SelectedIndex = 0;
        }

        private void select_File(Object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new()
            {
                Title = "Select The Required Excel Files",
                Filter = "Excel files (*.xlsx)|*.xlsx|Excel files (*.xlx)|*.xlx|All files (*.*)|*.*",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                Multiselect = true
            };
#if DEBUG
            openFileDialog.InitialDirectory = "F:\\001 - GEMICo\\002 - Mado\\Project 1\\003 - Documents\\Inbound-20230630T193845Z-001\\Inbound";
#endif

            if (openFileDialog.ShowDialog() == true)
            {
                Selected_Files.Text = String.Join(";\n", openFileDialog.FileNames);
                Selected_Files_List = openFileDialog.FileNames.ToList<String>();
            }
        }

        private void save_File(Object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog saveFileDialog = new();
            DialogResult saveFileDialogResult = saveFileDialog.ShowDialog();
            if (saveFileDialogResult.ToString().ToLower().Trim().Equals("ok"))
            {
                Destination = saveFileDialog.SelectedPath;
                Excel ex = new(Selected_Files_List);
                _ = ex.writeInbound(Destination, Joker_Value.Text);
            }
        }
    }
}
