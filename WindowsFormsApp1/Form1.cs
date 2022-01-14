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
                check.PaySystem = dataRows[i].Field<string>(0).ToString();
                check.PayMethod = dataRows[i].Field<string>(1).ToString();
                check.PayMethodInfo = dataRows[i].Field<string>(2).ToString();
                check.Date = dataRows[i].Field<DateTime?>(3).ToString();
                check.Sum = dataRows[i].Field<double>(4).ToString();
                check.Currency = dataRows[i].Field<string>(5).ToString();
                check.Login = dataRows[i].Field<string>(6).ToString();
                check.OrderNumber = dataRows[i].Field<string>(7).ToString();
                check.Statement = dataRows[i].Field<string>(8).ToString();
                check.PayDate = dataRows[i].Field<DateTime?>(9).ToString();
                check.WithdrawalDate = dataRows[i].Field<DateTime?>(10).ToString();
                check.ReturnDate = dataRows[i].Field<DateTime?>(11).ToString();
                check.AnswerCode = dataRows[i].Field<string>(12).ToString();
                check.IdOrderNumber = dataRows[i].Field<string>(13).ToString();
                check.CommissionSum = dataRows[i].Field<double>(14).ToString();
                check.ReturnSum = dataRows[i].Field<double>(15).ToString();
                check.OFDState = dataRows[i].Field<string>(16).ToString();
                check.OrderDescription = dataRows[i].Field<string>(17).ToString();
                list.Add(check);
                }

            #region
            //Check check = new Check(
            //    dataRows[i].Field<string>(0).ToString(), dataRows[i].Field<string>(1).ToString(),
            //    dataRows[i].Field<string>(2).ToString(), dataRows[i].Field<DateTime?>(3).ToString(),
            //    dataRows[i].Field<int>(4).ToString(), dataRows[i].Field<string>(5).ToString(),
            //    dataRows[i].Field<string>(6).ToString(), dataRows[i].Field<string>(7).ToString(),
            //    dataRows[i].Field<string>(8).ToString(), dataRows[i].Field<DateTime?>(9).ToString(),
            //    dataRows[i].Field<DateTime?>(10).ToString(), dataRows[i].Field<DateTime?>(11).ToString(),
            //    dataRows[i].Field<string>(12).ToString(), dataRows[i].Field<string>(13).ToString(),
            //    dataRows[i].Field<int>(14).ToString(), dataRows[i].Field<int>(15).ToString(),
            //    dataRows[i].Field<string>(16).ToString(), dataRows[i].Field<string>(17).ToString()
            //    );
            //foreach (var item in dataRows)
            //{
            //    Check check = new Check(
            //        item.Field<string>("Column0").ToString(), item.Field<string>("Column1").ToString(),
            //        item.Field<string>("Column2").ToString(), item.Field<DateTime?>("Column3").ToString(),
            //        item.Field<double>("Column4").ToString(), item.Field<string>("Column5").ToString(),
            //        item.Field<string>("Column6").ToString(), item.Field<string>("Column7").ToString(),
            //        item.Field<string>("Column8").ToString(), item.Field<DateTime?>("Column9").ToString(),
            //        item.Field<DateTime?>("Column10").ToString(), item.Field<DateTime?>("Column11").ToString(),
            //        item.Field<string>("Column12").ToString(), item.Field<string>("Column13").ToString(),
            //        item.Field<string>("Column14").ToString(), item.Field<string>("Column15").ToString(),
            //        item.Field<string>("Column16").ToString(), item.Field<string>("Column17").ToString()
            //        );
            //    list.Add(check);
            //}
            #endregion

            //перебор листа, сериализация каждого итема            
            int counter = 0;
                foreach (var item in list)
                {
                    XmlSerializer x = new XmlSerializer(item.GetType());

                    using (FileStream fstream = new FileStream($"C:/Users/Руслан/Desktop/test/doc_{counter}.xml", FileMode.OpenOrCreate))
                    {
                        x.Serialize(fstream, item);
                    }
                    counter++;
                }
                //using (FileStream fstream = new FileStream($"C:/Users/Руслан/Desktop/test/newXMLtable{i}.xml", FileMode.OpenOrCreate))
                //{
                //    byte[] array = Encoding.Default.GetBytes(dataSet.GetXml());
                //    fstream.Write(array, 0, array.Length);

                //    var a = dataSet.Tables[0].Select();
                //}
        }       

        
    }
}
