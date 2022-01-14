using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace WindowsFormsApp1
{
    public class Check
    {
        public Check(string paySystem, string payMethod, string payMethodInfo, 
            string date, string sum, string currency, string login, 
            string orderNumber, string statement, string payDate, string withdrawalDate, 
            string returnDate, string answerCode, string iDOrderNumber, string commissionSum, 
            string returnSum, string oFDState, string orderDescription)
        {
            PaySystem = paySystem;
            PayMethod = payMethod;
            PayMethodInfo = payMethodInfo;
            Date = date;
            Sum = sum;
            Currency = currency;
            Login = login;
            OrderNumber = orderNumber;
            Statement = statement;
            PayDate = payDate;
            WithdrawalDate = withdrawalDate;
            ReturnDate = returnDate;
            AnswerCode = answerCode;
            IdOrderNumber = iDOrderNumber;
            CommissionSum = commissionSum;
            ReturnSum = returnSum;
            OFDState = oFDState;
            OrderDescription = orderDescription;
        }
        public Check()
        {

        }

        [XmlElement]        public string PaySystem { get; set; }
        [XmlElement]        public string PayMethod { get; set; }
        [XmlElement]        public string PayMethodInfo { get; set; }
        [XmlElement]        public string Date { get; set; }
        [XmlElement]        public string Sum { get; set; }
        [XmlElement]        public string Currency { get; set; }
        [XmlElement]        public string Login { get; set; }
        [XmlElement]        public string OrderNumber { get; set; }
        [XmlElement]        public string Statement { get; set; }
        [XmlElement]        public string PayDate { get; set; }
        [XmlElement]        public string WithdrawalDate { get; set; }
        [XmlElement]        public string ReturnDate { get; set; }
        [XmlElement]        public string AnswerCode { get; set; }
        [XmlElement]        public string IdOrderNumber { get; set; }
        [XmlElement]        public string CommissionSum { get; set; }
        [XmlElement]        public string ReturnSum { get; set; }
        [XmlElement]        public string OFDState { get; set; }
        [XmlElement]        public string OrderDescription { get; set; }
    }
}
