using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using System.IO.Ports;
using Excel = Microsoft.Office.Interop.Excel;

namespace LAB
{
   public partial class Form1 : Form
    {
        static SerialPort COM;
        byte[] message = new byte[5];
        static bool inProcess = false;
        static Thread Rec;
        static bool zeroExists = false;
        static long step_now = 0;
        Button yes = new Button();
        Button no = new Button();
        Form dia = new Form();

        private byte[] convert(string x)
        {
            int p = 1, t = 0;
            for (int i = x.Count() - 1; i >= 0; i--)
            {
                t = t + ((int)(x[i]) - (int)('0')) * p;
                p = p * 10;
            }
            return BitConverter.GetBytes(t);
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            COM = new SerialPort(comboBox1.SelectedItem.ToString());
            if (comboBox1.SelectedItem != null && comboBox2.SelectedItem != null && comboBox3.SelectedItem != null && textBox1.Text != "")
                button1.Enabled = true;
            else button1.Enabled = false;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem != null && comboBox2.SelectedItem != null && comboBox3.SelectedItem != null && textBox1.Text != "")
                button1.Enabled = true;
            else button1.Enabled = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Rec = new Thread(()=>
            {
                Read(dataGridView1, textBox1, label7, label11);
            });
            try
            {
                if (!COM.IsOpen) COM.Open();
                COM.BaudRate = Convert.ToInt32(comboBox2.SelectedItem.ToString());
                COM.DtrEnable = true;
                message[1] = BitConverter.GetBytes(Convert.ToInt32(textBox1.Text))[0];
                message[2] = BitConverter.GetBytes(Convert.ToInt32(textBox1.Text))[1];
                message[3] = BitConverter.GetBytes(Convert.ToInt32(textBox1.Text))[2];
                message[4] = BitConverter.GetBytes(Convert.ToInt32(textBox1.Text))[3];
                label8.Visible = false;
                label9.Visible = false;
                label15.Visible = false;
            }
            catch (System.UnauthorizedAccessException)
            {
                label9.Visible = false;
                label8.Visible = false;
                label15.Visible = true;
                label7.Text = "";
                label11.Text = "";
            }
            catch (IOException)
            {
                label9.Visible = false;
                label15.Visible = false;
                label8.Visible = true;
                label7.Text = "";
                label11.Text = "";
            }
            catch (FormatException)
            {
                label15.Visible = false;
                label8.Visible = false;
                label9.Visible = true;
                label7.Text = "";
                label11.Text = "";
            }
            if (!label8.Visible && !label9.Visible && !label15.Visible)
            {
                for (int i = 0; i < 4; i++)
                {
                    if (comboBox3.Items[i] == comboBox3.SelectedItem)
                    {
                        message[0] = BitConverter.GetBytes(30 + i)[0];
                        COM.Write(message, 0, 5);
                        break;
                    }
                }
                inProcess = true;
                button1.Enabled = false;
                button4.Enabled = true;
                if (!Rec.IsAlive) Rec.Start();
            }
        }

        static void ToTable(DataGridView tab, int step, uint value)
        {
            if (tab.InvokeRequired) tab.Invoke((Action<DataGridView, int, uint>)ToTable, tab, step, value);
            else
            {
                if (step == 0)
                {
                    tab.Rows.Add(0, value);
                }
                else
                {
                    step_now += step;
                    tab.Rows.Add(step_now, value);
                }
            }
        }

        static void ToLabel(Label lab, string text)
        {
            if (lab.InvokeRequired) lab.Invoke((Action<Label, string>)ToLabel, lab, text);
            else lab.Text = text;
        }

        static void Read(DataGridView tab, TextBox steps, Label beg, Label end)
        {
            uint mode = 0;
            byte[] buff = new byte[5];
            ToLabel(beg, "");
            ToLabel(end, "");
            while (inProcess)
            {
                if (COM.BytesToRead == 5)
                {
                    COM.Read(buff, 0, 5);
                    mode = BitConverter.ToUInt32(buff, 1);

                    if (buff[0] == 0x60)
                    {
                        ToLabel(beg, mode.ToString());
                        if (!zeroExists)
                        {
                            ToTable(tab, 0, mode);
                            zeroExists = true;
                        }
                    }
                    else if (buff[0] == 0x61)
                    {
                        ToTable(tab, Convert.ToInt32(steps.Text), mode);
                        ToLabel(end, mode.ToString());
                    }
                    else if (buff[0] == 0x00)
                    {
                        inProcess = false;
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog saveTable = new SaveFileDialog();

                saveTable.Filter = "Excel files (*.xlsx)|*.xlsx";
                saveTable.FilterIndex = 1;
                saveTable.RestoreDirectory = true;

                if (saveTable.ShowDialog() == DialogResult.OK)
                {
                    Excel.Application app = new Excel.Application();
                    Excel.Workbook book = app.Workbooks.Add();
                    Excel.Worksheet sheet = book.ActiveSheet;

                    int i = 1;
                    while (sheet.Rows[1].Columns[i] == null && sheet.Rows[1].Columns[i + 1] == null) i++;
                    for (int j = 1; j <= dataGridView1.Rows.Count; j++)
                    {
                        sheet.Rows[j].Columns[i] = dataGridView1[0, j - 1].Value;
                        sheet.Rows[j].Columns[i + 1] = dataGridView1[1, j - 1].Value;
                    }
                    app.AlertBeforeOverwriting = false;

                    book.SaveAs(Filename: saveTable.FileName, Local: saveTable.InitialDirectory);
                    app.Quit();
                }
            }
            catch (System.Runtime.InteropServices.COMException) { }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem != null && comboBox2.SelectedItem != null && comboBox3.SelectedItem != null && textBox1.Text != "")
                button1.Enabled = true;
            else button1.Enabled = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            button4.Enabled = false;
            for (int i = 0; i < 5; i++) message[i] = 0x00;
            COM.Write(message, 0, 5);
            button1.Enabled = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            dia.MaximizeBox = false;
            dia.MinimizeBox = false;
            dia.Icon = this.Icon;
            dia.Text = "Очистка таблицы";
            dia.Width = 300;
            dia.Height = 150;
            dia.StartPosition = FormStartPosition.Manual;
            dia.Top = this.Top + this.Height / 2 - dia.Height / 2;
            dia.Left = this.Left + this.Width / 2 - dia.Width / 2;
            dia.FormBorderStyle = FormBorderStyle.FixedDialog;

            Label quest = new Label();

            quest.Width = 300;
            quest.Text = "Вы действительно хотите очистить таблицу?";
            quest.TextAlign = ContentAlignment.MiddleCenter;
            quest.Font = new Font(label1.Font.FontFamily, 10);
            quest.Top = 20;
            quest.Left = -10;

            Button yes = new Button();
            Button no = new Button();

            yes.Text = "Да";
            no.Text = "Нет!";
            yes.Font = new Font(label1.Font.FontFamily, 10);
            no.Font = new Font(label1.Font.FontFamily, 10);
            yes.Height = 35;
            no.Height = 35;
            yes.Location = new Point(dia.Width / 2 - yes.Width / 2 - 60, 60);
            no.Location = new Point(dia.Width / 2 - no.Width / 2 + 50, 60);

            dia.Controls.Add(quest);
            dia.Controls.Add(no);
            dia.Controls.Add(yes);

            yes.Click += yes_Click;
            no.Click += no_Click;
            dia.FormClosing += no_Click;

            this.Enabled = false;
            dia.Show();
        }

        private void yes_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            button3.Enabled = false;
            button2.Enabled = false;

            //dia.Close();
            dia.Visible = false;

            this.Enabled = true;
            this.Focus();
        }

        private void no_Click(object sender, EventArgs e)
        {

            //dia.Close();
            dia.Visible = false;

            this.Enabled = true;
            this.Focus();
        }

        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            if (dataGridView1.Rows.Count > 0 && inProcess == false)
            {
                button3.Enabled = true;
                button2.Enabled = true;
            }
        }
    }
}
