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

namespace DocumentsMerger
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private List<string> FilePaths { get; set; }
        public MainWindow()
        {
            InitializeComponent();           
        }

        private void MergeButton_Click(object sender, RoutedEventArgs e)
        {
            this.FilePaths = new List<string>(FolderSearch.GetAllFilePaths(this.PathTextBox.Text));
            Docs docs = new Docs(this.FilePaths);
            string output = @"C:\Users\1\Desktop\DocsMergerTest\merged.docx";
            //Merger.MergeOdt(docs, output);
            Merger.Merge(FilePaths.ToArray(), output, true, output);
            this.StatusLabel.Content = "Merging complete";
        }

        private void PathTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            
        }
    }
}
