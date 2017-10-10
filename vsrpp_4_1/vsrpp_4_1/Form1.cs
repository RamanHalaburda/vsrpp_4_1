using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;

namespace vsrpp_4_1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        Microsoft.Office.Interop.Word.Application wordapp;

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != String.Empty || textBox2.Text != String.Empty || textBox3.Text != String.Empty)
            {
                //wordapp = new Microsoft.Office.Interop.Word.Application();
                //wordapp.Visible = true;

                //Object saveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdPromptToSaveChanges;
                //Object originalFormat = Microsoft.Office.Interop.Word.WdOriginalFormat.wdWordDocument;
                //Object routeDocument = Type.Missing;
                //wordapp.Quit(ref saveChanges, ref originalFormat, ref routeDocument);
                //wordapp = null;

                Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
                Document doc = app.Documents.Open("testword.docx");
                object missing = System.Reflection.Missing.Value;
                //string s = "Hi";
                //Console.WriteLine(s);               
                doc.Content.Text = GenereteText(textBox1.Text, textBox2.Text, textBox3.Text);
                doc.Save();
                doc.Close(ref missing);

                app.Quit(ref missing);
            }
            else
            {
                MessageBox.Show("Заполните все поля!");
            }
        }

        private string GenereteText(string _receiver, string _email, string _sender)
        {
            return "\t"
                + "Уважаемый "
                + _receiver
                + "."
                + System.Environment.NewLine
                + System.Environment.NewLine
                + "\t"
                + "Я хотел бы узнать больше информации о размещении рекламы в вашей газете. "
                + "Пожалуйста, отправьте информации по адресу "
                + _email
                + "."
                + System.Environment.NewLine
                + System.Environment.NewLine
                + "\t"
                + "С уважением,"
                + System.Environment.NewLine
                + "\t"
                + _sender;
        }
    }
}
