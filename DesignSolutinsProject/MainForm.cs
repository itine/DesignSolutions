using MySql.Data.MySqlClient;
using RSDN;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DesignSolutinsProject
{
    public partial class MainForm : Form
    {
        MySqlConnection con = new MySqlConnection("server=localhost;userid=root;password=53344404;database=design_solutions");
        MySqlDataAdapter SDA = new MySqlDataAdapter();
        DataTable dbDataSet = new DataTable();
        BindingSource bSource = new BindingSource();
        public MainForm()
        {
            InitializeComponent();
            
        }
        private void button1_Click_1(object sender, EventArgs e)
        {
            CreateContract createContractForm = new CreateContract();
            MainForm mainForm = new MainForm();
            this.Visible = false;
            createContractForm.Show();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            try
            {
                con.Open();
                string Query = "select idcontract, docNumber as 'Номер договора', organizationName as 'Заказчик', inicial as 'ФИО', dateOfDocument as 'Дата', whatTheWork as 'Характеристика', totalCost as 'Итоговая сумма', daysForCompleted as 'Срок выполнения' from contract";
                MySqlCommand command = new MySqlCommand(Query, con);
                SDA.SelectCommand = command;
                SDA.Fill(dbDataSet);
                SDA.Update(dbDataSet);
                bSource.DataSource = dbDataSet;
                dataGridView1.DataSource = bSource;


                con.Close();
                dataGridView1.Columns["idcontract"].Visible = false;
                dataGridView1.Columns[2].Width = 120;
                dataGridView1.Columns[5].Width = 400;
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

        //поиск
        private void button2_Click(object sender, EventArgs e)
        {

            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                dataGridView1.CurrentCell = null;
                dataGridView1.Rows[i].Visible = false;
                dataGridView1.Rows[i].Selected = false;
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                        if (dataGridView1.Rows[i].Cells[j].Value.ToString().Contains(textBox1.Text))
                        {
                            dataGridView1.Rows[i].Visible = true;
                            break;
                        }
            }

        }
        
        
        //удаление
        private void button3_Click_1(object sender, EventArgs e)
        {
            string MyConnection2 = "server=localhost;userid=root;password=53344404;database=design_solutions";
            MySqlConnection MyConn2 = new MySqlConnection(MyConnection2);
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                DataGridViewRow dr = dataGridView1.Rows[i];
                if (dr.Selected == true)
                {
                    string id = Convert.ToString(dataGridView1.CurrentRow.Cells[0].Value);
                    try
                    {                      
                        string Query = "delete from contract where idcontract='" + id + "';";
                        MySqlCommand MyCommand2 = new MySqlCommand(Query, MyConn2);
                        MySqlDataReader MyReader2;
                        MyConn2.Open();
                        MyReader2 = MyCommand2.ExecuteReader();                        
                        MyConn2.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                    finally
                    {
                        MyConn2.Dispose();
                    }
                    dataGridView1.Rows.RemoveAt(i);
                }
            }
        }

        private void MainForm_ResizeEnd(object sender, EventArgs e)
        {
            dataGridView1.Size = this.Size;
            dataGridView1.Columns[5].Width = this.Size.Width / 2;
        }

        private void MainForm_Resize(object sender, EventArgs e)
        {
            dataGridView1.Size = this.Size;
            dataGridView1.Columns[5].Width = this.Size.Width / 2;
        }
        //вывод списка договоров
        private void button4_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                dataGridView1.CurrentCell = null;
                dataGridView1.Rows[i].Visible = false;
            }
            try
            {
                con.Open();
                string Query = "select idcontract, docNumber as 'Номер договора', organizationName as 'Заказчик', inicial as 'ФИО', dateOfDocument as 'Дата', whatTheWork as 'Характеристика', totalCost as 'Итоговая сумма', daysForCompleted as 'Срок выполнения' from contract";
                MySqlCommand command = new MySqlCommand(Query, con);
                SDA.SelectCommand = command;
                SDA.Fill(dbDataSet);
                bSource.DataSource = dbDataSet;
                dataGridView1.DataSource = bSource;
                SDA.Update(dbDataSet);

                con.Close();
                dataGridView1.Columns["idcontract"].Visible = false;
                dataGridView1.Columns[2].Width = 120;
                dataGridView1.Columns[5].Width = 400;
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
}
