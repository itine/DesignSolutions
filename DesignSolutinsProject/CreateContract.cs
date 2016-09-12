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
using Microsoft.CSharp;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Data.SqlClient;
using System.Data.Common;
using Excel = Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using DesignSolutinsProject;
using RSDN;
using System.IO;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.IO;
namespace DesignSolutinsProject
{
    public partial class CreateContract : Form
    {        
        public CreateContract()
        {
            InitializeComponent();
        }
        Word._Application application;
        Word._Document document;

        Excel._Application excelApp;
        Excel._Workbook workBook;

        Object missingObj = System.Reflection.Missing.Value;
        Object trueObj = true;
        Object falseObj = false;
        Object templatePathObj = " ";
        Object templatePathObj2 = " ";
        Object replaceParam1 = " ";
        Object replaceParam2 = " ";
        Object replaceParam12 = " ";
        Object replaceParam13 = " ";
        Object replaceParam14 = " ";
        Object replaceParam15 = " ";
        Object replaceParam17 = " ";
        Object replaceParam16 = " ";
        Object replaceParam18 = " ";
        Object replaceOrgName = " ";
        Object replaceParam7 = " ";
        Object replaceParam1111 = "";
        Object replaceParam5 = " ";
        Object replaceParam733 = "";
        Object replaceParam19 = " ";
        MySqlConnection con = new MySqlConnection("server=localhost;userid=root;password=53344404;database=design_solutions");
        MySqlDataAdapter SDA = new MySqlDataAdapter();
        System.Data.DataTable dbDataSet = new System.Data.DataTable();

        private void button4_Click(object sender, EventArgs e)
        {

            if (textBox8.Text == "") {
                MessageBox.Show("Заполните цену, используя разделитель \",\"");
                return;
            }

            if (textBox18.Text == "")
            {
                MessageBox.Show("Укажите предоплату");
                return;
            }
            else
            {
                application = new Word.Application();
                Object templatePathObj2 = "c:\\Работа\\#Архив Договоров\\DesignSolutinsProject\\DesignSolutinsProject\\Акт выполненных работ.doc";
                try
                {
                    document = application.Documents.Add(ref templatePathObj2, ref missingObj, ref missingObj, ref missingObj);

                }
                catch (Exception error)
                {
                    MessageBox.Show(error.ToString());
                    document.Close(ref falseObj, ref missingObj, ref missingObj);
                    application.Quit(ref missingObj, ref missingObj, ref missingObj);
                    document = null;
                    application = null;
                    throw error;
                }
                application.Visible = true;
                Object docNumber = "@@docNumber";
                replaceOrgName = textBox1.Text;
                Word.Range wordRange;
                object replaceTypeObj;
                replaceTypeObj = Word.WdReplace.wdReplaceAll;
                for (int i = 1; i <= document.Sections.Count; i++)
                {
                    wordRange = document.Sections[i].Range;
                    Word.Find wordFindObj = wordRange.Find;
                    object[] wordFindParameters = new object[15] { docNumber, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceOrgName, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                }
                object organizationName = "@@organizationName";
                replaceParam1 = textBox2.Text;
                Word.Range wordRange2;

                object replaceTypeObj2;
                replaceTypeObj2 = Word.WdReplace.wdReplaceAll;
                for (int i = 1; i <= document.Sections.Count; i++)
                {
                    wordRange2 = document.Sections[i].Range;
                    Word.Find wordFindObj = wordRange2.Find;
                    object[] wordFindParameters = new object[15] { organizationName, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam1, replaceTypeObj2, missingObj, missingObj, missingObj, missingObj };
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                }
                object city = "@@city";
                replaceParam2 = textBox3.Text;
                Word.Range wordRange3;
                object replaceTypeObj3;
                replaceTypeObj3 = Word.WdReplace.wdReplaceAll;
                for (int i = 1; i <= document.Sections.Count; i++)
                {
                    wordRange3 = document.Sections[i].Range;
                    Word.Find wordFindObj = wordRange3.Find;
                    object[] wordFindParameters = new object[15] { city, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam2, replaceTypeObj3, missingObj, missingObj, missingObj, missingObj };
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                }
                object stringDate = "@@stringDate";
                object replaceParam3 = DateTime.Now.ToShortDateString();
                Word.Range wordRange4;

                object replaceTypeObj4;
                replaceTypeObj4 = Word.WdReplace.wdReplaceAll;
                for (int i = 1; i <= document.Sections.Count; i++)
                {
                    wordRange4 = document.Sections[i].Range;
                    Word.Find wordFindObj = wordRange4.Find;
                    object[] wordFindParameters = new object[15] { stringDate, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam3, replaceTypeObj4, missingObj, missingObj, missingObj, missingObj };
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                }
                object client = "@@client";
                replaceParam5 = textBox4.Text;
                Word.Range wordRange6;

                object replaceTypeObj6;
                replaceTypeObj6 = Word.WdReplace.wdReplaceAll;
                for (int i = 1; i <= document.Sections.Count; i++)
                {
                    wordRange6 = document.Sections[i].Range;
                    Word.Find wordFindObj = wordRange6.Find;
                    object[] wordFindParameters = new object[15] { client, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam5, replaceTypeObj6, missingObj, missingObj, missingObj, missingObj };
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                }
                object whatTheWork = "@@whatTheWork";
                replaceParam7 = textBox10.Text;
                Word.Range wordRange8;

                object replaceTypeObj8;
                replaceTypeObj8 = Word.WdReplace.wdReplaceAll;
                for (int i = 1; i <= document.Sections.Count; i++)
                {
                    wordRange8 = document.Sections[i].Range;
                    Word.Find wordFindObj = wordRange8.Find;
                    object[] wordFindParameters = new object[15] { whatTheWork, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam7, replaceTypeObj8, missingObj, missingObj, missingObj, missingObj };
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                }
                object whatTheWork2 = "@@qwerty";
                replaceParam733 = textBox22.Text;
                Word.Range wordRange734;
                object replaceTypeObj734;
                replaceTypeObj734 = Word.WdReplace.wdReplaceAll;
                for (int i = 1; i <= document.Sections.Count; i++)
                {
                    wordRange734 = document.Sections[i].Range;
                    Word.Find wordFindObj = wordRange734.Find;
                    object[] wordFindParameters = new object[15] { whatTheWork2, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam733, replaceTypeObj734, missingObj, missingObj, missingObj, missingObj };
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                }
                double a = Double.Parse(textBox8.Text);
                int b = Int32.Parse(textBox18.Text);
                double res = a * b / 100;
                double myMoney = Math.Round(res, 2);
                object newCost = "@@newCost";
                object replaceParam9 = Convert.ToString(myMoney);
                Word.Range wordRange10;

                object replaceTypeObj10;
                replaceTypeObj10 = Word.WdReplace.wdReplaceAll;
                for (int i = 1; i <= document.Sections.Count; i++)
                {
                    wordRange10 = document.Sections[i].Range;
                    Word.Find wordFindObj = wordRange10.Find;
                    object[] wordFindParameters = new object[15] { newCost, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam9, replaceTypeObj10, missingObj, missingObj, missingObj, missingObj };
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                }

                string stringMoney = RusCurrency.Str(myMoney, "RUR");
                object myCost = "@@myCost";

                object replaceParam55 = stringMoney;
                Word.Range wordRange56;

                object replaceTypeObj56;
                replaceTypeObj56 = Word.WdReplace.wdReplaceAll;
                for (int i = 1; i <= document.Sections.Count; i++)
                {
                    wordRange56 = document.Sections[i].Range;
                    Word.Find wordFindObj = wordRange56.Find;
                    object[] wordFindParameters = new object[15] { myCost, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam55, replaceTypeObj56, missingObj, missingObj, missingObj, missingObj };
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                }

                object address = "@@address";
                replaceParam12 = textBox6.Text;
                Word.Range wordRange13;

                object replaceTypeObj13;
                replaceTypeObj13 = Word.WdReplace.wdReplaceAll;
                for (int i = 1; i <= document.Sections.Count; i++)
                {
                    wordRange13 = document.Sections[i].Range;
                    Word.Find wordFindObj = wordRange13.Find;
                    object[] wordFindParameters = new object[15] { address, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam12, replaceTypeObj13, missingObj, missingObj, missingObj, missingObj };
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                }

                object placeOfPayment = "@@placeOfPayment";
                replaceParam16 = textBox13.Text;
                Word.Range wordRange17;

                object replaceTypeObj17;
                replaceTypeObj17 = Word.WdReplace.wdReplaceAll;
                for (int i = 1; i <= document.Sections.Count; i++)
                {
                    wordRange17 = document.Sections[i].Range;
                    Word.Find wordFindObj = wordRange17.Find;
                    object[] wordFindParameters = new object[15] { placeOfPayment, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam16, replaceTypeObj17, missingObj, missingObj, missingObj, missingObj };
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                }
                object inn = "@@inn";
                replaceParam13 = textBox17.Text;
                Word.Range wordRange14;

                object replaceTypeObj14;
                replaceTypeObj14 = Word.WdReplace.wdReplaceAll;
                for (int i = 1; i <= document.Sections.Count; i++)
                {
                    wordRange14 = document.Sections[i].Range;
                    Word.Find wordFindObj = wordRange14.Find;
                    object[] wordFindParameters = new object[15] { inn, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam13, replaceTypeObj14, missingObj, missingObj, missingObj, missingObj };
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                }
                object kpp = "@@kpp";
                replaceParam14 = textBox16.Text;
                Word.Range wordRange15;

                object replaceTypeObj15;
                replaceTypeObj15 = Word.WdReplace.wdReplaceAll;
                for (int i = 1; i <= document.Sections.Count; i++)
                {
                    wordRange15 = document.Sections[i].Range;
                    Word.Find wordFindObj = wordRange15.Find;
                    object[] wordFindParameters = new object[15] { kpp, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam14, replaceTypeObj15, missingObj, missingObj, missingObj, missingObj };
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                }
                object rschet = "@@rschet";
                replaceParam15 = textBox15.Text;
                Word.Range wordRange16;

                object replaceTypeObj16;
                replaceTypeObj16 = Word.WdReplace.wdReplaceAll;
                for (int i = 1; i <= document.Sections.Count; i++)
                {
                    wordRange16 = document.Sections[i].Range;
                    Word.Find wordFindObj = wordRange16.Find;
                    object[] wordFindParameters = new object[15] { rschet, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam15, replaceTypeObj16, missingObj, missingObj, missingObj, missingObj };
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                }
                object kschet = "@@kschet";
                object replaceParam17 = textBox14.Text;
                Word.Range wordRange18;

                object replaceTypeObj18;
                replaceTypeObj18 = Word.WdReplace.wdReplaceAll;
                for (int i = 1; i <= document.Sections.Count; i++)
                {
                    wordRange18 = document.Sections[i].Range;
                    Word.Find wordFindObj = wordRange18.Find;
                    object[] wordFindParameters = new object[15] { kschet, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam17, replaceTypeObj18, missingObj, missingObj, missingObj, missingObj };
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                }

                object phoneNumber = "@@phoneNumber";
                replaceParam18 = textBox12.Text;
                Word.Range wordRange19;

                object replaceTypeObj19;
                replaceTypeObj19 = Word.WdReplace.wdReplaceAll;
                for (int i = 1; i <= document.Sections.Count; i++)
                {
                    wordRange19 = document.Sections[i].Range;
                    Word.Find wordFindObj = wordRange19.Find;
                    object[] wordFindParameters = new object[15] { phoneNumber, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam18, replaceTypeObj19, missingObj, missingObj, missingObj, missingObj };
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                }
                object pochta = "@@pochta";
                replaceParam1111 = textBox24.Text;
                Word.Range wordRange1112;

                object replaceTypeObj1112;
                replaceTypeObj1112 = Word.WdReplace.wdReplaceAll;
                for (int i = 1; i <= document.Sections.Count; i++)
                {
                    wordRange1112 = document.Sections[i].Range;
                    Word.Find wordFindObj = wordRange1112.Find;
                    object[] wordFindParameters = new object[15] { pochta, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam1111, replaceTypeObj1112, missingObj, missingObj, missingObj, missingObj };
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                }
                object clientInicial = "@@Inicial";
                replaceParam19 = textBox19.Text;
                Word.Range wordRange20;

                object replaceTypeObj20;
                replaceTypeObj20 = Word.WdReplace.wdReplaceAll;
                for (int i = 1; i <= document.Sections.Count; i++)
                {
                    wordRange20 = document.Sections[i].Range;
                    Word.Find wordFindObj = wordRange20.Find;
                    object[] wordFindParameters = new object[15] { clientInicial, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam19, replaceTypeObj20, missingObj, missingObj, missingObj, missingObj };
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                }
                object currentDate = "@@currentDate";
                object replaceParam20 = DateTime.Now.ToShortDateString();
                Word.Range wordRange21;

                object replaceTypeObj21;
                replaceTypeObj21 = Word.WdReplace.wdReplaceAll;
                for (int i = 1; i <= document.Sections.Count; i++)
                {
                    wordRange21 = document.Sections[i].Range;
                    Word.Find wordFindObj = wordRange21.Find;
                    object[] wordFindParameters = new object[15] { currentDate, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam20, replaceTypeObj21, missingObj, missingObj, missingObj, missingObj };
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                }
                object projectObject = "@@projectObject";
                object replaceParam200 = textBox11.Text;
                Word.Range wordRange201;

                object replaceTypeObj201;
                replaceTypeObj201 = Word.WdReplace.wdReplaceAll;
                for (int i = 1; i <= document.Sections.Count; i++)
                {
                    wordRange201 = document.Sections[i].Range;
                    Word.Find wordFindObj = wordRange201.Find;
                    object[] wordFindParameters = new object[15] { projectObject, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam200, replaceTypeObj201, missingObj, missingObj, missingObj, missingObj };
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                }
            }

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (textBox8.Text == "")
            {
                MessageBox.Show("Заполните цену, используя разделитель \",\"");
                return;
            }
            
            if (textBox18.Text == "")
            {
                MessageBox.Show("Укажите предоплату");
                return;
            }
            else
            {
                double total = Convert.ToDouble(textBox8.Text);
                string stringMoney = RusCurrency.Str(total, "RUR"); //число прописью
                application = new Word.Application();
               
                Object templatePathObj = "c:\\Работа\\#Архив Договоров\\DesignSolutinsProject\\DesignSolutinsProject\\Договор.doc";
                try
                {
                    document = application.Documents.Add(ref templatePathObj, ref missingObj, ref missingObj, ref missingObj);
                }
                catch (Exception error)
                {
                    MessageBox.Show(error.ToString());
                    document.Close(ref falseObj, ref missingObj, ref missingObj);
                    application.Quit(ref missingObj, ref missingObj, ref missingObj);
                    document = null;
                    application = null;
                    throw error;
                }
                application.Visible = true;
                try
                {
                    con.Open();
                    string dateNow = DateTime.Now.ToString("dd/MM/yyyy");
                    string Query = "insert into contract (docNumber,organizationName,inicial,projectObject,dateOfDocument,totalCost,daysForCompleted,orgCity, orgClient, reglamDoc, whatTheWork, prepayment, numberOfObject, orgAddress, orgInn, orgKpp, whatTheWorkForSchet,paymentAddress, phoneNumber, projectStudy, constructionTime, rSchet, kSchet) values (@docNumber,@organizationName,@inicial,@projectObject,@dateOfDocument,@totalCost,@daysForCompleted,@orgCity, @orgClient, @reglamDoc, @whatTheWork, @prepayment, @numberOfObject, @orgAddress, @orgInn, @orgKpp,@whatTheWorkForSchet,@paymentAddress, @phoneNumber, @projectStudy, @constructionTime, @rSchet, @kSchet)";
                    MySqlCommand command = new MySqlCommand(Query, con);
                    command.Parameters.AddWithValue("@dateOfDocument", dateNow);
                   

                    Object docNumber = "@@docNumber";
                    replaceOrgName = textBox1.Text;
                    Word.Range wordRange;

                    command.Parameters.AddWithValue("@docNumber", textBox1.Text);
                    object replaceTypeObj;
                    replaceTypeObj = Word.WdReplace.wdReplaceAll;
                    for (int i = 1; i <= document.Sections.Count; i++)
                    {
                        wordRange = document.Sections[i].Range;
                        Word.Find wordFindObj = wordRange.Find;
                        object[] wordFindParameters = new object[15] { docNumber, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceOrgName, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                        wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                    }
                    object organizationName = "@@organizationName";
                    replaceParam1 = textBox2.Text;
                    Word.Range wordRange2;
                    command.Parameters.AddWithValue("@organizationName", textBox2.Text);

                    object replaceTypeObj2;
                    replaceTypeObj2 = Word.WdReplace.wdReplaceAll;
                    for (int i = 1; i <= document.Sections.Count; i++)
                    {
                        wordRange2 = document.Sections[i].Range;
                        Word.Find wordFindObj = wordRange2.Find;
                        object[] wordFindParameters = new object[15] { organizationName, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam1, replaceTypeObj2, missingObj, missingObj, missingObj, missingObj };
                        wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                    }
                    object pochta = "@@pochta";
                    replaceParam1111 = textBox24.Text;
                    Word.Range wordRange1112;

                    object replaceTypeObj1112;
                    replaceTypeObj1112 = Word.WdReplace.wdReplaceAll;
                    for (int i = 1; i <= document.Sections.Count; i++)
                    {
                        wordRange1112 = document.Sections[i].Range;
                        Word.Find wordFindObj = wordRange1112.Find;
                        object[] wordFindParameters = new object[15] { pochta, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam1111, replaceTypeObj1112, missingObj, missingObj, missingObj, missingObj };
                        wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                    }
                    object myCost = "@@myCost";

                    object replaceParam55 = stringMoney;
                    Word.Range wordRange56;


                    object replaceTypeObj56;
                    replaceTypeObj56 = Word.WdReplace.wdReplaceAll;
                    for (int i = 1; i <= document.Sections.Count; i++)
                    {
                        wordRange56 = document.Sections[i].Range;
                        Word.Find wordFindObj = wordRange56.Find;
                        object[] wordFindParameters = new object[15] { myCost, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam55, replaceTypeObj56, missingObj, missingObj, missingObj, missingObj };
                        wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                    }

                    object city = "@@city";
                    replaceParam2 = textBox3.Text;
                    Word.Range wordRange3;
                    command.Parameters.AddWithValue("@orgCity", textBox3.Text);

                    object replaceTypeObj3;
                    replaceTypeObj3 = Word.WdReplace.wdReplaceAll;
                    for (int i = 1; i <= document.Sections.Count; i++)
                    {
                        wordRange3 = document.Sections[i].Range;
                        Word.Find wordFindObj = wordRange3.Find;
                        object[] wordFindParameters = new object[15] { city, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam2, replaceTypeObj3, missingObj, missingObj, missingObj, missingObj };
                        wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                    }
                    object currentDate = "@@currentDate";
                    object replaceParam3 = DateTime.Now.ToShortDateString();
                    Word.Range wordRange4;


                    object replaceTypeObj4;
                    replaceTypeObj4 = Word.WdReplace.wdReplaceAll;
                    for (int i = 1; i <= document.Sections.Count; i++)
                    {
                        wordRange4 = document.Sections[i].Range;
                        Word.Find wordFindObj = wordRange4.Find;
                        object[] wordFindParameters = new object[15] { currentDate, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam3, replaceTypeObj4, missingObj, missingObj, missingObj, missingObj };
                        wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                    }
                    object stringDate = "@@stringDate";
                    object replaceParam4 = DateTime.Now.ToLongDateString();
                    Word.Range wordRange5;


                    object replaceTypeObj5;
                    replaceTypeObj5 = Word.WdReplace.wdReplaceAll;
                    for (int i = 1; i <= document.Sections.Count; i++)
                    {
                        wordRange5 = document.Sections[i].Range;
                        Word.Find wordFindObj = wordRange5.Find;
                        object[] wordFindParameters = new object[15] { stringDate, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam4, replaceTypeObj5, missingObj, missingObj, missingObj, missingObj };
                        wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                    }
                    object client = "@@client";
                    replaceParam5 = textBox4.Text;
                    Word.Range wordRange6;
                    command.Parameters.AddWithValue("@orgClient", textBox4.Text);

                    object replaceTypeObj6;
                    replaceTypeObj6 = Word.WdReplace.wdReplaceAll;
                    for (int i = 1; i <= document.Sections.Count; i++)
                    {
                        wordRange6 = document.Sections[i].Range;
                        Word.Find wordFindObj = wordRange6.Find;
                        object[] wordFindParameters = new object[15] { client, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam5, replaceTypeObj6, missingObj, missingObj, missingObj, missingObj };
                        wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                    }
                    object type = "@@type";
                    object replaceParam6 = textBox5.Text;
                    Word.Range wordRange7;
                    command.Parameters.AddWithValue("@reglamDoc", textBox5.Text);

                    object replaceTypeObj7;
                    replaceTypeObj7 = Word.WdReplace.wdReplaceAll;
                    for (int i = 1; i <= document.Sections.Count; i++)
                    {
                        wordRange7 = document.Sections[i].Range;
                        Word.Find wordFindObj = wordRange7.Find;
                        object[] wordFindParameters = new object[15] { type, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam6, replaceTypeObj7, missingObj, missingObj, missingObj, missingObj };
                        wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                    }
                    object whatTheWork = "@@whatTheWork";
                    replaceParam7 = textBox10.Text;
                    Word.Range wordRange8;
                    command.Parameters.AddWithValue("@whatTheWork", textBox10.Text);

                    object replaceTypeObj8;
                    replaceTypeObj8 = Word.WdReplace.wdReplaceAll;
                    for (int i = 1; i <= document.Sections.Count; i++)
                    {
                        wordRange8 = document.Sections[i].Range;
                        Word.Find wordFindObj = wordRange8.Find;
                        object[] wordFindParameters = new object[15] { whatTheWork, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam7, replaceTypeObj8, missingObj, missingObj, missingObj, missingObj };
                        wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                    }
                    object whatTheWork2 = "@@qwerty";
                    replaceParam733 = textBox22.Text;
                    Word.Range wordRange734;
                    command.Parameters.AddWithValue("@whatTheWorkForSchet", (string)replaceParam733);

                    object replaceTypeObj734;
                    replaceTypeObj734 = Word.WdReplace.wdReplaceAll;
                    for (int i = 1; i <= document.Sections.Count; i++)
                    {
                        wordRange734 = document.Sections[i].Range;
                        Word.Find wordFindObj = wordRange734.Find;
                        object[] wordFindParameters = new object[15] { whatTheWork2, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam733, replaceTypeObj734, missingObj, missingObj, missingObj, missingObj };
                        wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                    }
                    object daysForCompleted = "@@daysForCompleted";
                    object replaceParam8 = textBox9.Text;
                    Word.Range wordRange9;
                    command.Parameters.AddWithValue("@daysForCompleted", textBox9.Text);

                    object replaceTypeObj9;
                    replaceTypeObj9 = Word.WdReplace.wdReplaceAll;
                    for (int i = 1; i <= document.Sections.Count; i++)
                    {
                        wordRange9 = document.Sections[i].Range;
                        Word.Find wordFindObj = wordRange9.Find;
                        object[] wordFindParameters = new object[15] { daysForCompleted, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam8, replaceTypeObj9, missingObj, missingObj, missingObj, missingObj };
                        wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                    }
                    object totalCost = "@@totalCost";
                    object replaceParam9 = textBox8.Text;
                    Word.Range wordRange10;
                    command.Parameters.AddWithValue("@totalCost", textBox8.Text);

                    object replaceTypeObj10;
                    replaceTypeObj10 = Word.WdReplace.wdReplaceAll;
                    for (int i = 1; i <= document.Sections.Count; i++)
                    {
                        wordRange10 = document.Sections[i].Range;
                        Word.Find wordFindObj = wordRange10.Find;
                        object[] wordFindParameters = new object[15] { totalCost, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam9, replaceTypeObj10, missingObj, missingObj, missingObj, missingObj };
                        wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                    }
                    object percent = "@@percent";
                    object replaceParam10 = textBox18.Text;
                    Word.Range wordRange11;

                    command.Parameters.AddWithValue("@prepayment", textBox18.Text);

                    object replaceTypeObj11;
                    replaceTypeObj11 = Word.WdReplace.wdReplaceAll;
                    for (int i = 1; i <= document.Sections.Count; i++)
                    {
                        wordRange11 = document.Sections[i].Range;
                        Word.Find wordFindObj = wordRange11.Find;
                        object[] wordFindParameters = new object[15] { percent, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam10, replaceTypeObj11, missingObj, missingObj, missingObj, missingObj };
                        wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                    }
                    object countOfCopies = "@@countOfCopies";
                    object replaceParam11 = textBox7.Text;
                    Word.Range wordRange12;
                    command.Parameters.AddWithValue("@numberOfObject", textBox7.Text);

                    object replaceTypeObj12;
                    replaceTypeObj12 = Word.WdReplace.wdReplaceAll;
                    for (int i = 1; i <= document.Sections.Count; i++)
                    {
                        wordRange12 = document.Sections[i].Range;
                        Word.Find wordFindObj = wordRange12.Find;
                        object[] wordFindParameters = new object[15] { countOfCopies, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam11, replaceTypeObj12, missingObj, missingObj, missingObj, missingObj };
                        wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                    }
                    object address = "@@address";
                    replaceParam12 = textBox6.Text;
                    Word.Range wordRange13;
                    command.Parameters.AddWithValue("@orgAddress", textBox6.Text);

                    object replaceTypeObj13;
                    replaceTypeObj13 = Word.WdReplace.wdReplaceAll;
                    for (int i = 1; i <= document.Sections.Count; i++)
                    {
                        wordRange13 = document.Sections[i].Range;
                        Word.Find wordFindObj = wordRange13.Find;
                        object[] wordFindParameters = new object[15] { address, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam12, replaceTypeObj13, missingObj, missingObj, missingObj, missingObj };
                        wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                    }
                    object inn = "@@inn";
                    replaceParam13 = textBox17.Text;
                    Word.Range wordRange14;
                    command.Parameters.AddWithValue("@orgInn", textBox17.Text);


                    object replaceTypeObj14;
                    replaceTypeObj14 = Word.WdReplace.wdReplaceAll;
                    for (int i = 1; i <= document.Sections.Count; i++)
                    {
                        wordRange14 = document.Sections[i].Range;
                        Word.Find wordFindObj = wordRange14.Find;
                        object[] wordFindParameters = new object[15] { inn, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam13, replaceTypeObj14, missingObj, missingObj, missingObj, missingObj };
                        wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                    }
                    object kpp = "@@kpp";
                    replaceParam14 = textBox16.Text;
                    Word.Range wordRange15;
                    command.Parameters.AddWithValue("@orgKpp", textBox16.Text);


                    object replaceTypeObj15;
                    replaceTypeObj15 = Word.WdReplace.wdReplaceAll;
                    for (int i = 1; i <= document.Sections.Count; i++)
                    {
                        wordRange15 = document.Sections[i].Range;
                        Word.Find wordFindObj = wordRange15.Find;
                        object[] wordFindParameters = new object[15] { kpp, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam14, replaceTypeObj15, missingObj, missingObj, missingObj, missingObj };
                        wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                    }
                    object rschet = "@@rschet";
                    replaceParam15 = textBox15.Text;
                    Word.Range wordRange16;
                    command.Parameters.AddWithValue("@rSchet", textBox15.Text);


                    object replaceTypeObj16;
                    replaceTypeObj16 = Word.WdReplace.wdReplaceAll;
                    for (int i = 1; i <= document.Sections.Count; i++)
                    {
                        wordRange16 = document.Sections[i].Range;
                        Word.Find wordFindObj = wordRange16.Find;
                        object[] wordFindParameters = new object[15] { rschet, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam15, replaceTypeObj16, missingObj, missingObj, missingObj, missingObj };
                        wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                    }
                    object kschet = "@@kschet";
                    object replaceParam17 = textBox14.Text;
                    Word.Range wordRange18;
                    command.Parameters.AddWithValue("@kSchet", textBox14.Text);


                    object replaceTypeObj18;
                    replaceTypeObj18 = Word.WdReplace.wdReplaceAll;
                    for (int i = 1; i <= document.Sections.Count; i++)
                    {
                        wordRange18 = document.Sections[i].Range;
                        Word.Find wordFindObj = wordRange18.Find;
                        object[] wordFindParameters = new object[15] { kschet, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam17, replaceTypeObj18, missingObj, missingObj, missingObj, missingObj };
                        wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                    }

                    object placeOfPayment = "@@placeOfPayment";
                    replaceParam16 = textBox13.Text;
                    Word.Range wordRange17;
                    command.Parameters.AddWithValue("@paymentAddress", textBox13.Text);


                    object replaceTypeObj17;
                    replaceTypeObj17 = Word.WdReplace.wdReplaceAll;
                    for (int i = 1; i <= document.Sections.Count; i++)
                    {
                        wordRange17 = document.Sections[i].Range;
                        Word.Find wordFindObj = wordRange17.Find;
                        object[] wordFindParameters = new object[15] { placeOfPayment, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam16, replaceTypeObj17, missingObj, missingObj, missingObj, missingObj };
                        wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                    }
                    object phoneNumber = "@@phoneNumber";
                    replaceParam18 = textBox12.Text;
                    Word.Range wordRange19;
                    command.Parameters.AddWithValue("@phoneNumber", textBox12.Text);


                    object replaceTypeObj19;
                    replaceTypeObj19 = Word.WdReplace.wdReplaceAll;
                    for (int i = 1; i <= document.Sections.Count; i++)
                    {
                        wordRange19 = document.Sections[i].Range;
                        Word.Find wordFindObj = wordRange19.Find;
                        object[] wordFindParameters = new object[15] { phoneNumber, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam18, replaceTypeObj19, missingObj, missingObj, missingObj, missingObj };
                        wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                    }
                    object clientInicial = "@@Inicial";
                    replaceParam19 = textBox19.Text;
                    Word.Range wordRange20;
                    command.Parameters.AddWithValue("@Inicial", textBox19.Text);

                    object replaceTypeObj20;
                    replaceTypeObj20 = Word.WdReplace.wdReplaceAll;
                    for (int i = 1; i <= document.Sections.Count; i++)
                    {
                        wordRange20 = document.Sections[i].Range;
                        Word.Find wordFindObj = wordRange20.Find;
                        object[] wordFindParameters = new object[15] { clientInicial, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam19, replaceTypeObj20, missingObj, missingObj, missingObj, missingObj };
                        wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                    }
                    object projectObject = "@@projectObject";
                    object replaceParam20 = textBox11.Text;
                    Word.Range wordRange21;
                    command.Parameters.AddWithValue("@projectObject", textBox11.Text);

                    object replaceTypeObj21;
                    replaceTypeObj21 = Word.WdReplace.wdReplaceAll;
                    for (int i = 1; i <= document.Sections.Count; i++)
                    {
                        wordRange21 = document.Sections[i].Range;
                        Word.Find wordFindObj = wordRange21.Find;
                        object[] wordFindParameters = new object[15] { projectObject, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam20, replaceTypeObj21, missingObj, missingObj, missingObj, missingObj };
                        wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                    }
                    object stage = "@@stage";
                    object replaceParam21 = textBox21.Text;
                    Word.Range wordRange22;
                    command.Parameters.AddWithValue("@projectStudy", textBox21.Text);


                    object replaceTypeObj22;
                    replaceTypeObj22 = Word.WdReplace.wdReplaceAll;
                    for (int i = 1; i <= document.Sections.Count; i++)
                    {
                        wordRange22 = document.Sections[i].Range;
                        Word.Find wordFindObj = wordRange22.Find;
                        object[] wordFindParameters = new object[15] { stage, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam21, replaceTypeObj22, missingObj, missingObj, missingObj, missingObj };
                        wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                    }
                    object constructionTime = "@@constructionTime";
                    object replaceParam22 = textBox20.Text;
                    Word.Range wordRange23;
                    command.Parameters.AddWithValue("@constructionTime", textBox20.Text);

                    object replaceTypeObj23;
                    replaceTypeObj23 = Word.WdReplace.wdReplaceAll;
                    for (int i = 1; i <= document.Sections.Count; i++)
                    {
                        wordRange23 = document.Sections[i].Range;
                        Word.Find wordFindObj = wordRange23.Find;
                        object[] wordFindParameters = new object[15] { constructionTime, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceParam22, replaceTypeObj23, missingObj, missingObj, missingObj, missingObj };
                        wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                    }
                    SDA.SelectCommand = command;
                    SDA.Fill(dbDataSet);

                    SDA.Update(dbDataSet);

                    con.Close();
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    con.Dispose();
                }
            }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (textBox8.Text == "")
            {
                MessageBox.Show("Заполните цену, используя разделитель \",\"");
                return;
            }
           
            if (textBox18.Text == "")
            {
                MessageBox.Show("Укажите предоплату");
                return;
            }
            else
            {
                excelApp = new Excel.Application();
                string filename = "c:\\Работа\\#Архив Договоров\\DesignSolutinsProject\\DesignSolutinsProject\\Счет.xls";

                double a = Double.Parse(textBox8.Text);
                int b = Int32.Parse(textBox18.Text);
                double res = a * b / 100;
                double myMoney = Math.Round(res, 2);

                excelApp.Workbooks.Open(
                                filename, // FileName
                                Type.Missing, // UpdateLinks
                                Type.Missing, //  ReadOnly
                                Type.Missing, // Format
                                Type.Missing, // Password
                                Type.Missing, // WriteResPassword
                                Type.Missing, // IgnoreReadOnlyRecommended
                                Type.Missing, // Origin
                                Type.Missing, // Delimiter
                                Type.Missing, // Editable
                                Type.Missing, //  Notify
                                Type.Missing, // Converter
                                Type.Missing, // AddToMru
                                Type.Missing, // Local
                                Type.Missing // CorruptLoad
                                );
                excelApp.Visible = true;

                Excel.Workbook book = excelApp.ActiveWorkbook;
                for (int i = 1; i <= book.Worksheets.Count; i++)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)book.Worksheets[i];
                    sheet.Cells[14, 2] = "СЧЕТ № " + textBox1.Text + " от " + DateTime.Now.ToLongDateString() + "";
                    sheet.Cells[17, 2] = "Плательщик:        " + textBox2.Text + " " + textBox3.Text + "";
                    sheet.Cells[18, 2] = "Грузополучатель: " + textBox2.Text + " " + textBox3.Text + "";
                    sheet.Cells[21, 3] = "Предоплата в размере "+ textBox18.Text + "%. По договору " + textBox1.Text + " от " + DateTime.Now.ToShortDateString() + " г. Проект " + textBox10.Text + textBox22.Text + " объекта, расположенного по адресу: " + textBox11.Text;
                    sheet.Cells[21, 7] = myMoney;
                    sheet.Cells[21, 8] = myMoney;
                    sheet.Cells[22, 8] = myMoney;
                    sheet.Cells[24, 8] = myMoney;
                    
                    sheet.Cells[26, 2] = "Всего наименований 1, на сумму " + myMoney;
                    sheet.Cells[27, 2] = RusCurrency.Str(myMoney, "RUR");
                    sheet.Cells[24, 8] = myMoney;
                }
                
                   
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MainForm mf = new MainForm();
            this.Visible = false;
            mf.Show();
        }
       
        private void textBox1_MouseClick(object sender, MouseEventArgs e)
        {
            textBox23.Text = "2016 / 0123";
        }

        private void textBox2_MouseClick(object sender, MouseEventArgs e)
        {
            textBox23.Text = "ООО \"Золотое дно\"";
        }

        private void textBox3_MouseClick(object sender, MouseEventArgs e)
        {
            textBox23.Text = "г. Покров";
        }

        private void textBox4_MouseClick(object sender, MouseEventArgs e)
        {
            textBox23.Text = "Директора Иванова Ивана Ивановича";
        }

        private void textBox5_MouseClick(object sender, MouseEventArgs e)
        {
            textBox23.Text = "устава";
        }

        private void textBox10_MouseClick(object sender, MouseEventArgs e)
        {
            textBox23.Text = "\"Техническое перевооружение внутриплощадочного надземного газопровода низкого давления Ø89 мм в части изменения положения относительно земли\" и \"Установка дополнительного газового оборудования\"";
        }

        private void textBox9_MouseClick(object sender, MouseEventArgs e)
        {
            textBox23.Text = "30";
        }

        private void textBox8_MouseClick(object sender, MouseEventArgs e)
        {
            textBox23.Text = "50000,00";
        }

        private void textBox18_MouseClick(object sender, MouseEventArgs e)
        {
            textBox23.Text = "50";
        }

        private void textBox7_MouseClick(object sender, MouseEventArgs e)
        {
            textBox23.Text = "3";
        }

        private void textBox6_MouseClick(object sender, MouseEventArgs e)
        {
            textBox23.Text = "601120, Владимирская обл., г. Покров, ул.Ленина, д. 181";
        }

        private void textBox11_MouseClick(object sender, MouseEventArgs e)
        {
            textBox23.Text = "\"Техническое перевооружение внутриплощадочного надземного газопровода низкого давления Ø89 мм в части изменения положения относительно земли\"  объекта, расположенного по адресу: 601122, Владимирская обл., Петушинский р - н, г.Покров, ул.Ленина, д. 181";
        }

        private void textBox13_MouseClick(object sender, MouseEventArgs e)
        {
            textBox23.Text = "Владимирский банк ФР №8611 г.Владимир";
        }

        private void textBox12_MouseClick(object sender, MouseEventArgs e)
        {
            textBox23.Text = "+7(777)-777-77-77";
        }

        private void textBox21_MouseClick(object sender, MouseEventArgs e)
        {
            textBox23.Text = "Одностадийное: Проектная документация";
        }

        private void textBox20_MouseClick(object sender, MouseEventArgs e)
        {
            textBox23.Text = "2016-2016";
        }

        private void textBox19_MouseClick(object sender, MouseEventArgs e)
        {
            textBox23.Text = "Иванов И.И.";
        }

        private void textBox17_MouseClick(object sender, MouseEventArgs e)
        {
            textBox23.Text = "5036003760";
        }

        private void textBox16_MouseClick(object sender, MouseEventArgs e)
        {
            textBox23.Text = "332101004";
        }

        private void textBox15_MouseClick(object sender, MouseEventArgs e)
        {
            textBox23.Text = "40702813910030102123";
        }

        private void textBox14_MouseClick(object sender, MouseEventArgs e)
        {
            textBox23.Text = "30101810020000000602";
        }

        private void textBox22_MouseClick(object sender, MouseEventArgs e)
        {
            textBox23.Text = "Предоплата в размере 50%. По договору 2016/0123 от 24.05.2016 г. Разработка проектной документации на \"Техническое перевооружение внутриплощадочного надземного газопровода низкого давления Ø89 мм в части изменения положения относительно земли\"  и \"Установка дополнительного газового оборудования\" объекта, расположенного по адресу: 601122, Владимирская обл., Петушинский р-н, г. Покров, ул. Ленина, д. 181";
        }

        private void CreateContract_Load(object sender, EventArgs e)
        {
            string Query = "select distinct organizationName from contract"; 
            using (con)
            using (MySqlCommand cmdDataBase = new MySqlCommand(Query, con))
            {
                try
                {
                    con.Open();
                    using (MySqlDataReader myReader = cmdDataBase.ExecuteReader())
                    {
                        while (myReader.Read())
                        {
                            string orgValue = (myReader.IsDBNull(myReader.GetOrdinal("organizationName")) ?
                                                string.Empty : myReader["organizationName"].ToString());
                            listBox1.Items.Add(orgValue);
                        }

                    }
                    con.Close();

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    con.Dispose();
                }
            }

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox22.Text = "";
            textBox22.Visible = false;
            string input = textBox1.Text.Trim();
            string text = listBox1.GetItemText(listBox1.SelectedItem);
            try
            {
                con.Open();
                string Query = "select distinct docNumber,whatTheWork,organizationName,inicial,projectObject,totalCost,daysForCompleted,orgCity, orgClient, reglamDoc, prepayment, numberOfObject, orgAddress, orgInn, orgKpp, paymentAddress, phoneNumber, projectStudy, constructionTime, rSchet, kSchet, whatTheWorkForSchet from contract where organizationName = @organizationName";
                MySqlCommand command = new MySqlCommand(Query, con);
                command.Parameters.AddWithValue("@organizationName", text);
                SDA.SelectCommand = command;
                System.Data.DataTable dt = new System.Data.DataTable();
                SDA.Fill(dt);

                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show(input + " is not exist.", "Not Exists");
                }
                else
                {
                    textBox24.Text = "Почтовый адрес заказчика";
                    textBox1.Text = dt.Rows[0][0] + "";
                    textBox10.Text = dt.Rows[0][1] + "";
                    textBox2.Text = dt.Rows[0][2] + "";
                    textBox19.Text = dt.Rows[0][3] + "";
                    textBox11.Text = dt.Rows[0][4] + "";
                    textBox8.Text = dt.Rows[0][5] + "";
                    textBox9.Text = dt.Rows[0][6] + "";
                    textBox3.Text = dt.Rows[0][7] + "";
                    textBox4.Text = dt.Rows[0][8] + "";
                    textBox5.Text = dt.Rows[0][9] + "";
                    textBox18.Text = dt.Rows[0][10] + "";
                    textBox7.Text = dt.Rows[0][11] + "";
                    textBox6.Text = dt.Rows[0][12] + "";
                    textBox17.Text = dt.Rows[0][13] + "";
                    textBox16.Text = dt.Rows[0][14] + "";
                    textBox13.Text = dt.Rows[0][15] + "";
                    textBox12.Text = dt.Rows[0][16] + "";
                    textBox21.Text = dt.Rows[0][17] + "";
                    textBox20.Text = dt.Rows[0][18] + "";
                    textBox15.Text = dt.Rows[0][19] + "";
                    textBox14.Text = dt.Rows[0][20] + "";
                    if ((string)dt.Rows[0][21] != "")
                    {
                        textBox22.Visible = true;
                        textBox22.Text = dt.Rows[0][21] + "";
                    }
                }
                con.Close();
            }
            catch (MySqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                con.Dispose();
            }
            
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
                listBox1.Visible = true;
            else
                listBox1.Visible = false;
        }

        private void textBox11_MouseClick_1(object sender, MouseEventArgs e)
        {
            textBox23.Text = "601122, Владимирская обл., Петушинский р-н, г. Покров, ул. Ленина, д. 181";

        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            if (textBox10.Text.Length == 174)
            {
                textBox22.Visible = true;
                textBox22.Focus();
            }
        }

        private void textBox22_MouseClick_1(object sender, MouseEventArgs e)
        {
            textBox23.Text = "Дополнительное окно для заполнения наименования работ";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Process.Start("charmap.exe");
        }

        private void textBox24_MouseClick(object sender, MouseEventArgs e)
        {
            textBox23.Text = "600009, г. Владимир ул. Суздальская д. 11, оф. 6";
        }
    }
}
