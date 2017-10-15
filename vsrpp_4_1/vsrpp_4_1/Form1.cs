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
            if (textBox1.Text != String.Empty && textBox2.Text != String.Empty && textBox3.Text != String.Empty)
            {
                CreateDocument();
            }
            else
            {
                MessageBox.Show("Заполните все поля!");
            }
        }

        private void CreateDocument()
        {
            try
            {
                Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();
                winword.ShowAnimation = false;
                winword.Visible = false;
                object missing = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word.Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);

                document.Content.Text = GenereteText(textBox1.Text, textBox2.Text, textBox3.Text);
                
                object currentFileName = @"e:\study\1_Study\BarsuEngineeringFaculty\ProjMVS2017\vsrpp_4_1\vsrpp_4_1\docx\"
                    + "(" + textBox1.Text + ") to (" + textBox1.Text + ").docx";

                document.SaveAs2(ref currentFileName);
                document.Close(ref missing, ref missing, ref missing);
                document = null;
                winword.Quit(ref missing, ref missing, ref missing);
                winword = null;
                MessageBox.Show("Document created successfully !");

                ClearFields();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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

        public void ClearFields()
        {
            textBox1.Text = textBox2.Text = textBox3.Text = String.Empty;
        }
    }
}
