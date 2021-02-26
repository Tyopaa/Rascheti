using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace RaschetPokupki
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        double width, height;
        double al = 15.5;
        double pl = 9.9;
        double sum = 0; 
        string check = "fdf";
        int i = 0;
        private void button1_Click(object sender, EventArgs e)
        {
            width = Convert.ToDouble(textBox1.Text);
            height = Convert.ToDouble(textBox2.Text);
            if(radioButton1.Checked)
            {
                check = radioButton1.Text;
                sum = width * height * al;
                label3.Text = "Размер: " + width.ToString("F2") + "x" + height.ToString("F2") + "см"
                + "\nМатериал: " + check + "\nСтоимость: " + sum + "₽";
            }
            if(radioButton2.Checked)
            {
                check = radioButton2.Text;
                sum = width * height * pl;
                label3.Text = "Размер: " + width.ToString("F2") + "x" + height.ToString("F2") + "см"
                + "\nМатериал: " + check + "\nСтоимость: " + sum + "₽";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string pathDocument = AppDomain.CurrentDomain.BaseDirectory + "Квитанция.docx";
            DocX document = DocX.Create(pathDocument);
            // вставляем параграф и передаём текст
            document.InsertParagraph("" +
                "ООО 'Уютный Дом' " +
                "\nДобро пожаловать " +
                "\nККМ 00075411 #3969 " +
                "\nИНН 1087746942040 " +
                "\nЭКЛЗ 3851495566 " +
                "\nЧек № " + i++ +
                "\n"+ DateTime.Now + " СИС.").
                     // устанавливаем шрифт
                     Font("Times New Roman").
                     // устанавливаем размер шрифта
                     FontSize(12).
                     // устанавливаем цвет
                     Color(Color.BlueViolet).
                     // выравниваем текст по центру
                     Alignment = Alignment.left;
            // создаём таблицу с 3 строками и 2 столбцами
            Table table = document.AddTable(6, 2);
            // располагаем таблицу по центру
            table.Alignment = Alignment.left;
            // меняем стандартный дизайн таблицы
            table.Design = TableDesign.TableGrid;
            // заполнение ячейки текстом
            table.Rows[0].Cells[0].Paragraphs[0].Append("Наименование товара").
                Bold().
                Color(Color.BlueViolet).
                FontSize(12).
                Alignment = Alignment.right;
            table.Rows[0].MergeCells(0, 1);
            table.Rows[1].Cells[0].Paragraphs[0].Append("Жалюзи").
                FontSize(12).
                Color(Color.BlueViolet).
                Alignment = Alignment.right;
            table.Rows[1].Cells[1].Paragraphs[0].Append(width.ToString("F2") + "x" + height.ToString("F2")).
                FontSize(12).
                Color(Color.BlueViolet).
                Alignment = Alignment.right;
            table.Rows[2].Cells[0].Paragraphs[0].Append("Материал").
                Color(Color.BlueViolet).
                FontSize(12).
                Alignment = Alignment.right;
            table.Rows[2].Cells[1].Paragraphs[0].Append(check).
                FontSize(12).
                Color(Color.BlueViolet).
                Alignment = Alignment.right;
            table.Rows[3].Cells[0].Paragraphs[0].Append("Итог").
                Color(Color.BlueViolet).
                FontSize(12).
                Alignment = Alignment.right;
            table.Rows[3].Cells[1].Paragraphs[0].Append("=" + sum).
                FontSize(12).
                Color(Color.BlueViolet).
                Alignment = Alignment.right;
            table.Rows[4].Cells[0].Paragraphs[0].Append("Сдача").
                Color(Color.BlueViolet).
                FontSize(12).
                Alignment = Alignment.right;
            table.Rows[4].Cells[1].Paragraphs[0].Append("=" + (sum - sum)).
                FontSize(12).
                Color(Color.BlueViolet).
                Alignment = Alignment.right;
            table.Rows[5].Cells[0].Paragraphs[0].Append("Сумма итого").
                Color(Color.BlueViolet).
                FontSize(12).
                Alignment = Alignment.right;
            table.Rows[5].Cells[1].Paragraphs[0].Append("=" + sum).
                Color(Color.BlueViolet).
                FontSize(12).
                Alignment = Alignment.right;
            // создаём параграф и вставляем таблицу
            document.InsertParagraph().InsertTableAfterSelf(table);
            document.InsertParagraph();
            document.InsertParagraph("************************").
                Color(Color.BlueViolet).
                FontSize(12).
                Alignment = Alignment.left;
            document.InsertParagraph("00003751#059705").
                Color(Color.BlueViolet).
                FontSize(12).
                Alignment = Alignment.left;
            // сохраняем документ
            document.Save();
        }
    }
}

