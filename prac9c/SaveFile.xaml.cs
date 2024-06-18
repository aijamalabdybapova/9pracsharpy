using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Windows;

namespace prac9c
{
    public partial class SaveFile : Window
    {
        private string _filePath;

        public SaveFile(string filePath)
        {
            InitializeComponent();
            _filePath = filePath;

            MinHeight = 200;
            MinWidth = 400;
        }



        private void Button_Click(object sender, RoutedEventArgs e)
        {

            try
            {
                MailMessage mail = new MailMessage(From.Text, To.Text, Subject.Text, "письмо");

                if (File.Exists(_filePath))
                {
                    mail.Attachments.Add(new Attachment(_filePath));
                }
                else
                {
                    MessageBox.Show("Выбран неверный файл!");
                    return;
                }

                SmtpClient smtp = new SmtpClient();
                smtp.EnableSsl = true;

                if (From.Text.Contains("@mail.ru"))
                {
                    smtp.Host = "smtp.mail.ru";
                    smtp.Port = 587;
                }
                else if (From.Text.Contains("@yandex.ru"))
                {
                    smtp.Host = "smtp.yandex.ru";
                    smtp.Port = 25;
                }
                else if (From.Text.Contains("@rambler.ru"))
                {
                    smtp.Host = "smtp.rambler.ru";
                    smtp.Port = 25;
                }
                else if (From.Text.Contains("@gmail.com"))
                {
                    smtp.Host = "smtp.gmail.com";
                    smtp.Port = 587;
                }
                else
                {
                    MessageBox.Show("Неверный хост или домен");
                    return;
                }

                smtp.Credentials = new NetworkCredential(From.Text, Pass.Text);

                try
                {
                    smtp.Send(mail);
                    MessageBox.Show("Письмо успешно отправлено!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка при отправке письма: " + ex.Message);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Вы не ввели данные для отправки! " + ex.Message);
            }
        }
    }
}