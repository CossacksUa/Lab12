using System;
using System.IO;
using System.Windows.Forms;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace lab12
{
    public partial class Form1 : Form
    {
        private string _selectedTemplatePath;
        private string _selectedDocumentPath;

        public Form1()
        {
            InitializeComponent();
            LoadTemplates();
        }

        private void LoadTemplates()
        {
            string templatesDirectory = @"C:\Users\Monolit\Desktop\ООП\ЛР\Lab12\Templates";
            var templates = Directory.GetFiles(templatesDirectory, "*.docx");

            foreach (var template in templates)
            {
                listBox1.Items.Add(Path.GetFileName(template));
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem != null)
            {
                _selectedTemplatePath = Path.Combine(@"C:\Users\Monolit\Desktop\ООП\ЛР\Lab12\Templates", listBox1.SelectedItem.ToString());
                string outputPath = Path.Combine(@"C:\Users\Monolit\Desktop\ООП\ЛР\Lab12\GeneratedDocuments", $"GeneratedDocument_{DateTime.Now:yyyyMMddHHmmss}.docx");

                string placeholder1Text = textBox1.Text;
                string placeholder2Text = textBox2.Text;
                string placeholder5Text = textBox5.Text;
                string placeholder6Text = textBox6.Text;
                string placeholder7Text = textBox7.Text;
                string placeholder8Text = textBox8.Text;
                string placeholder9Text = textBox9.Text;



                GenerateDocument(_selectedTemplatePath, outputPath, placeholder1Text, placeholder2Text, placeholder5Text, placeholder6Text, placeholder7Text, placeholder8Text, placeholder9Text);

                MessageBox.Show($"Документ створено: {outputPath}");
            }
            else
            {
                MessageBox.Show("Будь ласка, оберіть шаблон.");
            }
        }

        private void GenerateDocument(string templatePath, string outputPath, string placeholder1Text, string placeholder2Text, string placeholder5Text, string placeholder6Text, string placeholder7Text, string placeholder8Text, string placeholder9Text)
        {
            File.Copy(templatePath, outputPath, true);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(outputPath, true))
            {
                ReplacePlaceholderText(doc.MainDocumentPart.Document, "{1}", placeholder1Text);
                ReplacePlaceholderText(doc.MainDocumentPart.Document, "{2}", placeholder2Text);
                ReplacePlaceholderText(doc.MainDocumentPart.Document, "{3}", placeholder5Text);
                ReplacePlaceholderText(doc.MainDocumentPart.Document, "{Adress}", placeholder6Text);
                ReplacePlaceholderText(doc.MainDocumentPart.Document, "{Nomber}", placeholder7Text);
                ReplacePlaceholderText(doc.MainDocumentPart.Document, "{Email}", placeholder8Text);
                ReplacePlaceholderText(doc.MainDocumentPart.Document, "{Web}", placeholder9Text);
                doc.MainDocumentPart.Document.Save();
            }
        }

        private void ReplacePlaceholderText(Document document, string placeholder, string newText)
        {
            if (string.IsNullOrEmpty(placeholder) || newText == null)
            {
                MessageBox.Show("Текст для пошуку не може бути пустим.");
                return;
            }

            foreach (var text in document.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
            {
                if (text.Text.Contains(placeholder))
                {
                    text.Text = text.Text.Replace(placeholder, newText);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Word Documents|*.docx";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    _selectedDocumentPath = openFileDialog.FileName;
                    MessageBox.Show($"Документ обрано: {_selectedDocumentPath}");
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (_selectedDocumentPath == null)
            {
                MessageBox.Show("Будь ласка, оберіть документ для пошуку та заміни.");
                return;
            }

            string searchText = textBox3.Text;
            string replaceText = textBox4.Text;

            if (string.IsNullOrEmpty(searchText))
            {
                MessageBox.Show("Текст для пошуку не може бути пустим.");
                return;
            }

            SearchAndReplaceInDocument(_selectedDocumentPath, searchText, replaceText);
            MessageBox.Show("Заміна завершена.");
        }

        private void SearchAndReplaceInDocument(string filePath, string searchText, string replaceText)
        {
            string outputPath = Path.Combine(Path.GetDirectoryName(filePath), $"Modified_{Path.GetFileName(filePath)}");
            File.Copy(filePath, outputPath, true);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(outputPath, true))
            {
                ReplacePlaceholderText(doc.MainDocumentPart.Document, searchText, replaceText);
                doc.MainDocumentPart.Document.Save();
            }
        }
    }
}
