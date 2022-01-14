using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private DataSet dataSet;
        private string folderPath = "C:/Users/Руслан/Desktop/test/";
        private string defaultFileName = "default";
        
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string fileName = openFileDialog1.FileName;
                LoadFromExcel(fileName);
                button2.Enabled = true;
            }
        }
        private void LoadFromExcel(string fileName)
        {
            DataSet ds = new DataSet();
            try
            {
                using (var stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        do
                        {
                            while (reader.Read())
                            {
                            }
                        }
                        while (reader.NextResult());
                        ds = reader.AsDataSet();                        
                        dataSet = ds;
                        textBox1.Text = fileName;
                    }
                }
            }
            catch (IOException ex)
            {
                MessageBox.Show("Нужно закрыть файл: " + fileName);
                return;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            //создание массива из рядов таблицы, создание коллекции, содержащей экземпляры класса Чек
            DataRow[] dataRows = dataSet.Tables[0].Select();            
            List<Check> list = new List<Check>();

            //перебор массива, создание экземпляров Чек, добавление в лист 
            for (int i = 1; i < dataRows.Length; i++)
            {
                Check check = new Check();

                try { check.PaySystem = dataRows[i].Field<string>(0).ToString();}
                catch (Exception) { check.PaySystem = DBNull.Value.ToString(); }

                try { check.PayMethod = dataRows[i].Field<string>(1).ToString(); }
                catch (Exception) { check.PayMethod = DBNull.Value.ToString(); }

                try { check.PayMethodInfo = dataRows[i].Field<string>(2).ToString(); }
                catch (Exception) { check.PayMethodInfo = DBNull.Value.ToString(); }

                try { check.Date = dataRows[i].Field<DateTime?>(3).ToString(); }
                catch (Exception) { check.Date = DBNull.Value.ToString(); }

                try { check.Sum = dataRows[i].Field<double>(4).ToString(); }
                catch (Exception) { check.Sum = DBNull.Value.ToString(); }

                try { check.Currency = dataRows[i].Field<string>(5).ToString(); }
                catch (Exception) { check.Currency = DBNull.Value.ToString(); }

                try { check.Login = dataRows[i].Field<string>(6).ToString(); }
                catch (Exception) { check.Login = DBNull.Value.ToString(); }

                try { check.OrderNumber = dataRows[i].Field<string>(7).ToString(); }
                catch (Exception) { check.OrderNumber = DBNull.Value.ToString(); }

                try { check.Statement = dataRows[i].Field<string>(8).ToString(); }
                catch (Exception) { check.Statement = DBNull.Value.ToString(); }

                try { check.PayDate = dataRows[i].Field<DateTime?>(9).ToString(); }
                catch (Exception) { check.PayDate = DBNull.Value.ToString(); }

                try { check.WithdrawalDate = dataRows[i].Field<DateTime?>(10).ToString(); }
                catch (Exception) { check.WithdrawalDate = DBNull.Value.ToString(); }

                try { check.ReturnDate = dataRows[i].Field<DateTime?>(11).ToString(); }
                catch (Exception) { check.ReturnDate = DBNull.Value.ToString(); }

                try { check.AnswerCode = dataRows[i].Field<string>(12).ToString(); }
                catch (Exception) { check.AnswerCode = DBNull.Value.ToString(); }

                try { check.IdOrderNumber = dataRows[i].Field<string>(13).ToString(); }
                catch (Exception) { check.IdOrderNumber = DBNull.Value.ToString(); }

                try { check.CommissionSum = dataRows[i].Field<double>(14).ToString(); }
                catch (Exception) { check.CommissionSum = DBNull.Value.ToString(); }

                try { check.ReturnSum = dataRows[i].Field<double>(15).ToString(); }
                catch (Exception) { check.ReturnSum = DBNull.Value.ToString(); }

                try { check.OFDState = dataRows[i].Field<string>(16).ToString(); }
                catch (Exception) { check.OFDState = DBNull.Value.ToString(); }

                try { check.OrderDescription = dataRows[i].Field<string>(17).ToString(); }
                catch (Exception) { check.OrderDescription = DBNull.Value.ToString(); }                
                
                list.Add(check);
            }            

            //перебор листа, сериализация каждого итема            
            int counter = 0;

            foreach (var item in list)
            {
                XmlSerializer x = new XmlSerializer(item.GetType());

                using (FileStream fstream = new FileStream($"{folderPath}{defaultFileName}{counter}.xml", FileMode.OpenOrCreate))
                {
                    x.Serialize(fstream, item);
                }
                counter++;
            }
            MessageBox.Show($"Конвертация прошла успешно! \n Файлы сохранены в {folderPath}");
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            folderPath = $"{textBox4.Text}/";
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            defaultFileName = textBox6.Text;
        }

        
    }
}
