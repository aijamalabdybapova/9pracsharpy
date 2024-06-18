using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using Xceed.Words.NET;  // Добавлено для работы с DocX

namespace prac9c
{
    public partial class WORDLOAD : Window
    {
        public WORDLOAD()
        {
            InitializeComponent();
        }

        void LoadRtfFile(string _fileName)
        {
            if (File.Exists(_fileName))
            {
                TextRange range = new TextRange(MyRtb.Document.ContentStart, MyRtb.Document.ContentEnd);
                using (FileStream fStream = new FileStream(_fileName, FileMode.Open))
                {
                    range.Load(fStream, DataFormats.Rtf);
                }
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.Filters.Add(new CommonFileDialogFilter("RTF Files", ".rtf"));
            CommonFileDialogResult result = dialog.ShowDialog();
            if (result == CommonFileDialogResult.Ok)
            {
                string fileName = dialog.FileName;
                LoadRtfFile(fileName);
            }
        }

        void SaveRtbFileAsDocx(string _fileName)
        {
            TextRange range = new TextRange(MyRtb.Document.ContentStart, MyRtb.Document.ContentEnd);
            string plainText = range.Text;

            using (var document = DocX.Create(_fileName))
            {
                document.InsertParagraph(plainText);
                document.Save();
            }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            CommonSaveFileDialog dialog = new CommonSaveFileDialog();
            dialog.Filters.Add(new CommonFileDialogFilter("Word Document", "*.docx"));
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                string filePath = dialog.FileName;
                SaveRtbFileAsDocx(filePath);
                MessageBox.Show("Файл сохранен: " + filePath);
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.Filters.Add(new CommonFileDialogFilter("Word Files", "*.doc;*.docx"));
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                string filePath = dialog.FileName;
                SaveFile send = new SaveFile(filePath);
                send.ShowDialog();
            }
        }
    }
}