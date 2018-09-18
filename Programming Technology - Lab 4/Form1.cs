using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;


namespace Programming_Technology___Lab_4
{
    public partial class Form1 : Form
    {
        Word.Application application;
        Word.Document document;
        string fileName = null;
        string savePatch = null;


        public Form1()
        {
            InitializeComponent();            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            application = new Word.Application();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "doc files (*.doc)|*.doc|docx files (*.docx)|*.docx";
            

            try
            {
                    if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    fileName = openFileDialog1.FileName;
                }
           
                document = application.Documents.Add(fileName);
                application.Visible = true;
            }
            catch (Exception error)
            {
                MessageBox.Show("Произошла ошибка при попытке открыть файл");
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            document = application.Documents.Add();
            application.Visible = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            application.Visible = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            application.Visible = true;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                document.Close();
                //application.Quit();
                document = null;
                //application = null;
            }
            catch (Exception error)
            {
                MessageBox.Show("Произошла ошибка при попытке закрыть файл");
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "doc files (*.doc)|*.doc|docx files (*.docx)|*.docx";
            try
            {
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    savePatch = saveFileDialog1.FileName;
                }

                document.SaveAs(savePatch);
            }
            catch (Exception error)
            {
                MessageBox.Show("Произошла ошибка при попытке сохранения");
            }
        }
    }
}
