    using Microsoft.WindowsAPICodePack.Dialogs;
    using System;
    using System.Collections.Generic;
    using System.IO;
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
    using System.Windows.Shapes;
    using Spire.Doc;

    namespace prac9c
    {
        /// <summary>
        /// Логика взаимодействия для WORD.xaml
        /// </summary>
        public partial class WORD : Window
        {
        public string path;
            public WORD()
            {
                InitializeComponent();
                this.path = path;
            }
       
            
        private bool SaveFile()
        {
            CommonSaveFileDialog dialog = new CommonSaveFileDialog();
            dialog.Filters.Add(new CommonFileDialogFilter("Word файл", "*.docx"));

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                string fileName = dialog.FileName;

                if (!fileName.EndsWith(".docx"))
                {
                    fileName += ".docx";
                }

                var range = new TextRange(MyRtb.Document.ContentStart, MyRtb.Document.ContentEnd);
                using (var fs = new FileStream(fileName, FileMode.Create))
                {
                    range.Save(fs, DataFormats.Rtf);
                }

                Document document = new Document();
                document.LoadFromFile(fileName);
                document.SaveToFile(fileName, FileFormat.Docx);

                path = fileName;

                return true;
            }

            return false;
        }
        
            private void Button_Click(object sender, RoutedEventArgs e)

            {
            SaveFile();
            }

            private void Button_Click_1(object sender, RoutedEventArgs e)
            {
                CommonOpenFileDialog dialog = new CommonOpenFileDialog();
                dialog.Filters.Add(new CommonFileDialogFilter("Word Files", "*.doc;*.docx"));
                CommonFileDialogResult result = dialog.ShowDialog();
                if (result == CommonFileDialogResult.Ok)
                {
                    string filePath = dialog.FileName;
                    SaveFile send = new SaveFile(filePath);
                    send.ShowDialog();
                }
            }
        }
    }
