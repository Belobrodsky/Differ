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
using System.Windows.Forms.DataVisualization.Charting;



//псевдонимы
using SD = System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using SmallPertubation;
using Cursor = System.Windows.Forms.Cursor;

namespace Converter
{
    public struct OneRec
    {
        public double dt;
        public double value;
    }

    public partial class Form1 : Form
    {
        private System.Collections.ArrayList customers = new System.Collections.ArrayList();
        private MyVirtualClass customerInEdit;
        private int rowInEdit = -1;
        private bool rowScopeCommit = true;

        private System.Collections.ArrayList customers1 = new System.Collections.ArrayList();
        private MyVirtualClass customerInEdit1;
        private int rowInEdit1 = -1;
        private bool rowScopeCommit1 = true;

        private System.Collections.ArrayList customers2 = new System.Collections.ArrayList();
        private MyVirtualClass customerInEdit2;
        private int rowInEdit2 = -1;
        private bool rowScopeCommit2 = true;

        private System.Collections.ArrayList customers3 = new System.Collections.ArrayList();
        private MyVirtualClass customerInEdit3;
        private int rowInEdit3 = -1;
        private bool rowScopeCommit3 = true;

        private System.Collections.ArrayList customers4 = new System.Collections.ArrayList();
        private MyVirtualClass customerInEdit4;
        private int rowInEdit4 = -1;
        private bool rowScopeCommit4 = true;

        private System.Collections.ArrayList customers5 = new System.Collections.ArrayList();
        private MyVirtualClass customerInEdit5;
        private int rowInEdit5 = -1;
        private bool rowScopeCommit5 = true;

        private System.Collections.ArrayList customers6 = new System.Collections.ArrayList();
        private MyVirtualClass customerInEdit6;
        private int rowInEdit6 = -1;
        private bool rowScopeCommit6 = true;

        public Form1()
        {
            InitializeComponent();




            chart1.Series["Series1"].Color = Color.Red;
            chart1.Series["Series2"].Color = Color.Blue;
            chart1.Series["Series3"].Color = Color.Lime;
            chart1.Series["Series4"].Color = Color.Aqua;

            chart1.ChartAreas[0].AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dash;
            chart1.ChartAreas[0].AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dash;
            chart1.ChartAreas[0].AxisX.MajorGrid.LineColor = Color.DarkGray;
            chart1.ChartAreas[0].AxisY.MajorGrid.LineColor = Color.DarkGray;

            chart1.ChartAreas[0].AxisX.MinorTickMark.Enabled = true;
            chart1.ChartAreas[0].AxisX.MinorTickMark.LineDashStyle = ChartDashStyle.Dash;
            chart1.ChartAreas[0].AxisX.MinorTickMark.LineColor = Color.DarkGray;

            //chart1.ChartAreas[0].AxisY.Interval = 0.006;
            chart1.ChartAreas[0].AxisY.MinorTickMark.Enabled = true;
            chart1.ChartAreas[0].AxisY.MinorTickMark.LineDashStyle = ChartDashStyle.Dash;
            chart1.ChartAreas[0].AxisY.MinorTickMark.LineColor = Color.DarkGray;

            //   chart1.ChartAreas[0].AxisX.Minimum = 0;
            //    chart1.ChartAreas[0].AxisX.Interval =1000;


            //  chart1.ChartAreas[0].AxisY2.Interval =10;
            chart1.ChartAreas[0].AxisY2.MinorTickMark.Enabled = true;
            chart1.ChartAreas[0].AxisY2.MinorTickMark.LineDashStyle = ChartDashStyle.Dash;
            chart1.ChartAreas[0].AxisY2.MinorTickMark.LineColor = Color.DarkGray;

            chart1.Legends["Legend1"].BorderColor = Color.Black;

        }

        private MyListOfSensors MyAllSensors = new MyListOfSensors();


        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {


        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            MINnew.Clear();
            MAXnew.Clear();

            for (int i = 0; i < MaXes.Count; i++)
            {

                //textBox1.Text = dataGridView2.Rows[i].Cells[2].Value.ToString();
                //    MAXnew.Add(Convert.ToDouble(textBox1.Text));

                //  textBox1.Text = dataGridView2.Rows[i].Cells[1].Value.ToString();
                //  MINnew.Add(Convert.ToDouble(textBox1.Text));
            } //
            //textBox1.Text = openFileDialog1.FileName;
        }

        private void очиститьToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            //  открытьToolStripMenuItem.Enabled = true;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            //открытьToolStripMenuItem.Enabled = true;
        }

        private void очиститьГрафикToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Form2 main = this.Owner as Form2;
            checkedListBox1.Items.Clear();
            MyAllSensors.Clear();
            AllLEGENDS.Clear();

            //открытьToolStripMenuItem.Enabled = true;

            //  NumberSeries = 0;
            for (int i = 0; i < chart1.Series.Count; i++)
            {
                chart1.Series[i].Points.Clear();
                chart1.Series[i].LegendText = "";
            }

            //убирает все галочки с чеклистбоксов
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, false);
            }
            chart1.Titles.Clear();

            for (int i = 1; i < chart1.Series.Count - 2; i++)
            {
                chart1.Series["Series" + i].IsVisibleInLegend = false;
                //   chart1.Series["Series" + i].Points.Clear();
                //  chart1.Series["Series" + i].Points.Clear();
            }
            //button1.Enabled = false;
            chart1.Annotations.Clear();
            //очиститьГрафикToolStripMenuItem.Enabled = false;
            //   поменятьФорматОсиXНаВременнойToolStripMenuItem.Enabled = false;
            //  checkBox1.Enabled = false;
            //checkBox4.Enabled = false;
            // checkBox5.Enabled = false;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void checkedListBox1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                checkedListBox1.ContextMenuStrip = contextMenuStrip1;
                checkedListBox1.SelectedIndex = checkedListBox1.IndexFromPoint(e.X, e.Y);
            }
        }

        private void добатьбToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {
        }



        private List<string> AllLEGENDS = new List<string>();

        //    private double[] indexTimelist;
        private bool flagAxis = true;
        private List<double> _timeList = new List<double>();
        private List<int> _indexList = new List<int>();

        public double Min;
        public double Max;

        private void добавитьНаОсьXToolStripMenuItem_Click(object sender, EventArgs e)
        {
            flagAxis = true;
            Graph.CreateLine(MyAllSensors, checkedListBox1.Text, chart1, flagAxis);
            Min = chart1.ChartAreas[0].AxisY.Minimum;
            Max = chart1.ChartAreas[0].AxisY.Maximum;

            добавитьНаДополнительнуюОсьYОтВремениToolStripMenuItem.Enabled = true;
            _timeList.Clear();
            _indexList.Clear();
            for (int i = 0; i < MyAllSensors[0].MyListRecordsForOneKKS.Count; i++)
            {
                _timeList.Add(MyAllSensors[0].MyListRecordsForOneKKS[i].ValueTimeForDAT);
                _indexList.Add(i);
            }
        }

        private double prosent;

        private List<double> MaXes = new List<double>();
        private List<double> MiNes = new List<double>();

        private void Form1_Load(object sender, EventArgs e)
        {
            добавитьНаОсьXToolStripMenuItem.Enabled = false;
            добавитьНаДополнительнуюОсьYОтВремениToolStripMenuItem.Enabled = false;
            button6.Enabled = false;
            textBox15.Text = 0.ToString();

            comboBox6.Items.Add("Подбор");
            comboBox6.Items.Add("3");
            comboBox6.Items.Add("4");
            comboBox6.Items.Add("5");

            comboBox6.Text = comboBox6.Items[0].ToString();

            dataGridView5.Columns.Add("Время", "Время");
            dataGridView5.Columns.Add("Jотн", "Jотн");
            dataGridView5.Columns.Add("Jкор", "Jкор");
            dataGridView5.Columns.Add("Rэкс", "Rэкс");
            dataGridView5.Columns.Add("Rрасч", "Rрасч");
            dataGridView5.Columns.Add("dH12", "dH12");
            dataGridView5.Columns.Add("F", "F");

            tabPage5.Text = "Все данные";
            tabPage6.Text = "Данные для поиска ПЭ и J(отн)";
            tabPage7.Text = "Расчет";
            tabPage8.Text = "ПЭ";
            tabPage9.Text = "Данные из расчета dp/dH";

            chart1.ChartAreas[0].CursorX.IsUserEnabled = true;
            chart1.ChartAreas[0].AxisX.ScrollBar.IsPositionedInside = true;
            chart1.ChartAreas[0].AxisX.ScrollBar.Size = 20;
            chart1.ChartAreas[0].AxisY.ScrollBar.Size = 20;
            chart1.ChartAreas[0].CursorY.IsUserEnabled = true;
            chart1.ChartAreas[0].CursorY.LineColor = Color.Blue;
            chart1.ChartAreas[0].CursorX.LineColor = Color.Blue;
            chart1.ChartAreas[0].CursorY.LineWidth = 2;
            chart1.ChartAreas[0].CursorX.LineWidth = 2;
            chart1.ChartAreas[0].CursorY.LineDashStyle = ChartDashStyle.Dash;
            chart1.ChartAreas[0].CursorX.LineDashStyle = ChartDashStyle.Dash;
            chart1.ChartAreas[0].CursorY.Interval = 0.0000001;
            chart1.ChartAreas[0].CursorX.Interval = 0.0000001;
            chart1.ChartAreas[0].CursorY.SelectionColor = Color.Blue;


            button8.Enabled = false;

            textBox3.BackColor = Color.Red;
            textBox4.BackColor = Color.Red;
            tabPage1.Text = "Зона обработки";
            tabPage2.Text = "Результаты";
            button6.Visible = true;
            button6.BackColor = Color.Yellow;
            button7.Visible = false;
            tabPage3.Text = "Параметры";
            tabPage4.Text = "Данные";
            // button6.Enabled = false;
            chart1.ChartAreas[0].BackColor = Color.Gainsboro;

            for (int i = 0; i < chart1.Series.Count; i++)
                chart1.Series[i].IsVisibleInLegend = false;

            chart1.ChartAreas[0].AxisX.TitleFont = new System.Drawing.Font("Times New Roman", 14, FontStyle.Regular);
            chart1.ChartAreas[0].AxisY.TitleFont = new System.Drawing.Font("Times New Roman", 14, FontStyle.Regular);

            dataGridView2.Columns.Add("\u03B1(H) \u03B2/см", "\u03B1(H) \u03B2/см");
            dataGridView2.Columns.Add("По пичкам \u03B2/см", "По пичкам \u03B2/см");
            dataGridView2.Columns.Add("TAU", "TAU");

            this.dataGridView1.VirtualMode = true;
            this.dataGridView4.VirtualMode = true;

            DataGridViewTextBoxColumn companyNameColumn = new DataGridViewTextBoxColumn();
            companyNameColumn.HeaderText = "Время";
            companyNameColumn.Name = "Время";
            this.dataGridView1.Columns.Add(companyNameColumn);

            DataGridViewTextBoxColumn companyNameColumn1 = new DataGridViewTextBoxColumn();
            companyNameColumn1.HeaderText = "Значение";
            companyNameColumn1.Name = "Значение";
            this.dataGridView1.Columns.Add(companyNameColumn1);


            tableLayoutPanel10.ColumnStyles[3].Width = 0;
            tableLayoutPanel10.ColumnStyles[2].Width = 2F;

            dataGridView1.Columns[0].Width = 200;
            dataGridView1.Columns[1].Width = 200;

            label24.Text = "T, \u2103";
            label1.Text = "\u03C1, %";
        }

        private void dataGridView1_CellValueNeeded(object sender,
            System.Windows.Forms.DataGridViewCellValueEventArgs e)
        {
            //this.dataGridView1.RowCount = 1;
            // If this is the row for new records, no values are needed.
            if (e.RowIndex == this.dataGridView1.RowCount - 1) return;
            //   if (e.RowIndex == this.dataGridView3.RowCount - 1) return;

            MyVirtualClass customerTmp = null;

            // Store a reference to the Customer object for the row being painted.
            if (e.RowIndex == rowInEdit)
            {
                customerTmp = this.customerInEdit;
            }
            else
            {
                customerTmp = (MyVirtualClass) this.customers[e.RowIndex];
            }

            // Set the cell value to paint using the Customer object retrieved.
            switch (this.dataGridView1.Columns[e.ColumnIndex].Name)
            {
                case "Время":
                    e.Value = customerTmp.time;
                    break;

                case "Значение":
                    e.Value = customerTmp.value;
                    break;
            }
        }

        private void dataGridView4_CellValueNeeded(object sender,
            System.Windows.Forms.DataGridViewCellValueEventArgs e)
        {
            //this.dataGridView1.RowCount = 1;
            // If this is the row for new records, no values are needed.
            if (e.RowIndex == this.dataGridView4.RowCount - 1) return;
            //   if (e.RowIndex == this.dataGridView3.RowCount - 1) return;

            MyVirtualClass customerTmp = null;

            // Store a reference to the Customer object for the row being painted.
            if (e.RowIndex == rowInEdit1)
            {
                customerTmp = this.customerInEdit1;
            }
            else
            {
                customerTmp = (MyVirtualClass) this.customers1[e.RowIndex];
            }

            // Set the cell value to paint using the Customer object retrieved.
            switch (this.dataGridView4.Columns[e.ColumnIndex].Name)
            {
                case "Время":
                    e.Value = customerTmp.value;
                    break;
            }
        }



        private void dataGridView41_CellValueNeeded(object sender,
            System.Windows.Forms.DataGridViewCellValueEventArgs e)
        {
            //this.dataGridView1.RowCount = 1;
            // If this is the row for new records, no values are needed.
            if (e.RowIndex == this.dataGridView4.RowCount - 1) return;
            //   if (e.RowIndex == this.dataGridView3.RowCount - 1) return;

            MyVirtualClass customerTmp = null;

            // Store a reference to the Customer object for the row being painted.
            if (e.RowIndex == rowInEdit2)
            {
                customerTmp = this.customerInEdit2;
            }
            else
            {
                customerTmp = (MyVirtualClass) this.customers2[e.RowIndex];
            }

            // Set the cell value to paint using the Customer object retrieved.
            switch (this.dataGridView4.Columns[e.ColumnIndex].Name)
            {
                case "Ток(эксперимент), А":
                    e.Value = customerTmp.value;
                    break;
            }
        }

        private void dataGridView42_CellValueNeeded(object sender,
            System.Windows.Forms.DataGridViewCellValueEventArgs e)
        {
            //this.dataGridView1.RowCount = 1;
            // If this is the row for new records, no values are needed.
            if (e.RowIndex == this.dataGridView4.RowCount - 1) return;
            //   if (e.RowIndex == this.dataGridView3.RowCount - 1) return;

            MyVirtualClass customerTmp = null;

            // Store a reference to the Customer object for the row being painted.
            if (e.RowIndex == rowInEdit3)
            {
                customerTmp = this.customerInEdit3;
            }
            else
            {
                customerTmp = (MyVirtualClass) this.customers3[e.RowIndex];
            }

            // Set the cell value to paint using the Customer object retrieved.
            switch (this.dataGridView4.Columns[e.ColumnIndex].Name)
            {
                case "\u03C1, %":
                    e.Value = customerTmp.value;
                    break;
            }
        }

        private void dataGridView43_CellValueNeeded(object sender,
            System.Windows.Forms.DataGridViewCellValueEventArgs e)
        {
            //this.dataGridView1.RowCount = 1;
            // If this is the row for new records, no values are needed.
            if (e.RowIndex == this.dataGridView4.RowCount - 1) return;
            //   if (e.RowIndex == this.dataGridView3.RowCount - 1) return;

            MyVirtualClass customerTmp = null;

            // Store a reference to the Customer object for the row being painted.
            if (e.RowIndex == rowInEdit4)
            {
                customerTmp = this.customerInEdit4;
            }
            else
            {
                customerTmp = (MyVirtualClass) this.customers4[e.RowIndex];
            }

            // Set the cell value to paint using the Customer object retrieved.
            switch (this.dataGridView4.Columns[e.ColumnIndex].Name)
            {
                case "N, МВт":
                    e.Value = customerTmp.value;
                    break;
            }
        }

        private void dataGridView44_CellValueNeeded(object sender,
            System.Windows.Forms.DataGridViewCellValueEventArgs e)
        {
            //this.dataGridView1.RowCount = 1;
            // If this is the row for new records, no values are needed.
            if (e.RowIndex == this.dataGridView4.RowCount - 1) return;
            //   if (e.RowIndex == this.dataGridView3.RowCount - 1) return;

            MyVirtualClass customerTmp = null;

            // Store a reference to the Customer object for the row being painted.
            if (e.RowIndex == rowInEdit5)
            {
                customerTmp = this.customerInEdit5;
            }
            else
            {
                customerTmp = (MyVirtualClass) this.customers5[e.RowIndex];
            }

            // Set the cell value to paint using the Customer object retrieved.
            switch (this.dataGridView4.Columns[e.ColumnIndex].Name)
            {
                case "H12, см":
                    e.Value = customerTmp.value;
                    break;
            }
        }

        private void dataGridView45_CellValueNeeded(object sender,
            System.Windows.Forms.DataGridViewCellValueEventArgs e)
        {
            //this.dataGridView1.RowCount = 1;
            // If this is the row for new records, no values are needed.
            if (e.RowIndex == this.dataGridView4.RowCount - 1) return;
            //   if (e.RowIndex == this.dataGridView3.RowCount - 1) return;

            MyVirtualClass customerTmp = null;

            // Store a reference to the Customer object for the row being painted.
            if (e.RowIndex == rowInEdit6)
            {
                customerTmp = this.customerInEdit6;
            }
            else
            {
                customerTmp = (MyVirtualClass) this.customers6[e.RowIndex];
            }

            // Set the cell value to paint using the Customer object retrieved.
            switch (this.dataGridView4.Columns[e.ColumnIndex].Name)
            {
                case "T, \u2103":
                    e.Value = customerTmp.value;
                    break;
            }
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.Visible = false;
                //   Sencors myOneKKS = new Sencors();
                Sensors myOneKKS = new Sensors();
                //e.Node.Parent.Index
                //  List<double> MyParametr = new List<double>();
                //    MessageBox.Show(MyListSensors.Count.ToString());
                //    MessageBox.Show(MyListSensors[0].Count.ToString() + " " + MyListSensors[1].Count.ToString());
                MyListOfSensors N = new MyListOfSensors();
                myOneKKS = N.getOneKKSByIndex(checkedListBox1.SelectedIndex, MyAllSensors);

                //     List<double> MyParametr = new List<double>();



                //     myOneKKS = MyListSensors[1].getOneKKSByIndex(treeView1.SelectedNode.Index);

                //  myOneKKS = MyAllSensors.getOneKKSByIndex(treeView1.SelectedNode.Index);
                //         MessageBox.Show(myOneKKS.KKS_Name);
                this.dataGridView1.CellValueNeeded += new
                    DataGridViewCellValueEventHandler(dataGridView1_CellValueNeeded);

                dataGridView1.Rows.Clear();
                customers.Clear();

                for (int i = 0; i < myOneKKS.MyListRecordsForOneKKS.Count; i++)
                {
                    this.customers.Add(
                        new MyVirtualClass(myOneKKS.MyListRecordsForOneKKS[i].DateTime.ToString("HH:mm:ss.fff"),
                            myOneKKS.MyListRecordsForOneKKS[i].Value.ToString()));

                }
                if (this.dataGridView1.RowCount == 0)
                {
                    this.dataGridView1.RowCount = 1;
                    //  .ToString("dd.MM.yy HH:mm:ss.fff")
                }

                this.dataGridView1.RowCount = myOneKKS.MyListRecordsForOneKKS.Count;
                dataGridView1.Visible = true;
            }
            catch
            {
                //  MessageBox.Show(ex0.Message);
            }
        }

        private void chart1_MouseDown(object sender, MouseEventArgs e)
        {


        }

        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                chart1.SaveImage(saveFileDialog1.FileName + ".TIFF", System.Drawing.Imaging.ImageFormat.Jpeg);
            }
        }

        private void параметрыГрафикаToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void добавитьНаВспомагательнуюОсьYToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void добавитьНаОсновнуюОсьYToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        //    int NumberSeries = 0;

        private void убратьПодписиКривыхToolStripMenuItem_Click(object sender, EventArgs e)
        {
            chart1.Legends["Legend1"].Enabled = false;
        }

        private TextAnnotation text1 = new TextAnnotation();

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void добавитьНаВспомогательнуюОсьXToolStripMenuItem_Click(object sender, EventArgs e)
        {


        }

        private void добавитьНаВспомогательнуюОсьXToolStripMenuItem_MouseDown(object sender, MouseEventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void инструкцияПоПрименениюToolStripMenuItem_Click(object sender, EventArgs e)
        {


        }

        //Word.Application word = new Word.Application();
        private void справкаToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void очиститьВсеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Restart();
        }

        private void jToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void осьY1ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void осьY2ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void осьY3ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void осьY4ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void форматВремениДоСекундToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {



        }

        private void поменятьФорматОсиXНаВременнойToolStripMenuItem_Click(object sender, EventArgs e)
        {
            chart1.Series[0].XValueType = ChartValueType.DateTime;
            //очиститьГрафикToolStripMenuItem.Enabled = false;
            //  поменятьФорматОсиXНаВременнойToolStripMenuItem.Enabled = false;
            //добавитьНаОсьYToolStripMenuItem.Enabled = false;
            //checkBox4.Enabled = true;
        }

        private void checkBox4_CheckedChanged_1(object sender, EventArgs e)
        {

            chart1.ChartAreas[0].CursorX.IsUserEnabled = true;
            chart1.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart1.ChartAreas[0].AxisX.ScaleView.Zoomable = true;
            chart1.ChartAreas[0].AxisX.ScrollBar.IsPositionedInside = true;



        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {



        }

        public int NumberPoints;

        private void button2_Click_1(object sender, EventArgs e)
        {
            //   NumberPoints = (int)int.Parse(textBox2.Text);
            // MessageBox.Show(NumberPoints.ToString());

            List<double> TERT = new List<double>();
            TERT.Clear();

            double sum = 0;
            double aver = 0;

            for (int ii = 0; ii < 1; ii++)
            {
                sum = 0;
                aver = 0;
                for (int j = 0; j < MyAllSensors.Count; j++)
                {
                    if (chart1.Series[ii].LegendText == MyAllSensors[j].KKS_Name &&
                        (int) chart1.ChartAreas[0].CursorX.Position > 0)
                    {
                        for (int i = 0; i < NumberPoints; i++)
                        {
                            sum = sum +
                                  MyAllSensors[j].MyListRecordsForOneKKS[(int) chart1.ChartAreas[0].CursorX.Position - i
                                      ].Value;
                        }
                        aver = sum/NumberPoints;
                        TERT.Add(aver);

                        //dataGridView3.Rows[0].Cells[ii + 1].Value = TERT[ii];
                        //  MessageBox.Show(TERT[ii].ToString());
                    }
                }
            }
        }

        private void copySelectedRowsToClipboard(DataGridView dgv)
        {
            dgv.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            Clipboard.Clear();
            if (dgv.GetClipboardContent() != null)
            {
                Clipboard.SetDataObject(dgv.GetClipboardContent());
                Clipboard.GetData(DataFormats.Text);
                IDataObject dt = Clipboard.GetDataObject();
                if (dt.GetDataPresent(typeof (string)))
                {
                    string tb = (string) (dt.GetData(typeof (string)));
                    Encoding encoding = Encoding.GetEncoding(1251);
                    byte[] dataStr = encoding.GetBytes(tb);
                    Clipboard.SetDataObject(encoding.GetString(dataStr));
                }
            }
            dgv.ClipboardCopyMode = DataGridViewClipboardCopyMode.Disable;
        }

        private void dataGridView2_MouseDown(object sender, MouseEventArgs e)
        {



        }

        private void wxToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.C)
                if (((DataGridView) sender).SelectedCells.Count > 0)
                {
                    copySelectedRowsToClipboard((DataGridView) sender);
                }
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            //   if (e.Control && e.KeyCode == Keys.C)
            //    if (((DataGridView)sender).SelectedCells.Count > 0)
            //  {
            //     copySelectedRowsToClipboard((DataGridView)sender);
            //  }
        }

        private List<double> MAXnew = new List<double>();
        private List<double> MINnew = new List<double>();

        private void button3_Click(object sender, EventArgs e)
        {
            //     textBox1.Text = dataGridView2.Rows[6].Cells[2].Value.ToString();
            MINnew.Clear();
            MAXnew.Clear();
            for (int i = 0; i < MaXes.Count; i++)
            {

                //textBox1.Text = dataGridView2.Rows[i].Cells[2].Value.ToString();
                // MAXnew.Add(Convert.ToDouble(textBox1.Text));

                //    textBox1.Text = dataGridView2.Rows[i].Cells[1].Value.ToString();
                //   MINnew.Add(Convert.ToDouble(textBox1.Text));
            }
        }

        private void перестрToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int indexGrafic;
            for (int j = 0; j < AllLEGENDS.Count; j++)
            {
                if ((string) checkedListBox1.Text.Split('\t')[0] == AllLEGENDS[j])
                {
                    indexGrafic = j;
                    chart1.Series[indexGrafic].Points.Clear();
                    {
                        for (int k = 0; k < MyAllSensors.Count; k++)
                        {
                            if (AllLEGENDS[j] == MyAllSensors[k].KKS_Name)
                            {
                                for (int i = 0; i < MyAllSensors[k].MyListRecordsForOneKKS.Count; i++)
                                {
                                    prosent = ((MyAllSensors[k].MyListRecordsForOneKKS[i].Value - MINnew[k])/
                                               (MAXnew[k] - MINnew[k]))*100;
                                    chart1.Series[indexGrafic].Points.AddXY(i, prosent);
                                }
                                chart1.Series[indexGrafic].IsVisibleInLegend = true;
                                chart1.Series[indexGrafic].LegendText = MyAllSensors[k].KKS_Name;
                                chart1.Series[indexGrafic].YAxisType = AxisType.Secondary;
                            }
                        }
                    }
                }
            }
        }

        private void показатьДанныеToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.C)
                if (((DataGridView) sender).SelectedCells.Count > 0)
                {
                    copySelectedRowsToClipboard((DataGridView) sender);
                }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            chart1.ChartAreas[0].CursorX.Position--;
            //MessageBox.Show(chart1.ChartAreas[0].CursorX.Position.ToString());


            for (int ii = 0; ii < 1; ii++)
            {
//колличество нарисованных сириосов тоесть колличество выводимых параметров
                //   TochkaDannih.Columns.Add(chart1.Series[ii].LegendText);
                for (int j = 0; j < MyAllSensors.Count; j++)
                {
                    if (chart1.Series[ii].LegendText == MyAllSensors[j].KKS_Name)
                    {
                        //  DateTime WindowsTime = new DateTime(1970, 1, 1).AddSeconds(allSensors[j].MyListRecordsForOneKKS[(int)chart1.ChartAreas[0].CursorX.Position].DateTime);
                        //   dataGridView3.Rows[0].Cells[0].Value = WindowsTime.ToString("HH:mm:ss");


                        textBox3.Text =
                            MyAllSensors[j].MyListRecordsForOneKKS[(int) chart1.ChartAreas[0].CursorX.Position]
                                .ValueTimeForDAT.ToString();
                        textBox4.Text =
                            MyAllSensors[j].MyListRecordsForOneKKS[(int) chart1.ChartAreas[0].CursorX.Position].DateTime
                                .ToString("HH:mm:ss");
                    }
                }
            }
            //for (int ii = 0; ii < NumberSeriesNew; ii++)
            //{//колличество нарисованных сириосов тоесть колличество выводимых параметров
            //    TochkaDannih.Columns.Add(chart1.Series[ii].LegendText);
            //    for (int j = 0; j < allSensors.Count; j++)
            //    {
            //        if (chart1.Series[ii].LegendText == allSensors[j].KksName)
            //        {
            //            DateTime WindowsTime = new DateTime(1970, 1, 1).AddSeconds(allSensors[j].values[(int)chart1.ChartAreas[0].CursorX.Position].dt);

            //            dataGridView3.Rows[0].Cells[0].Value = WindowsTime.ToString("HH:mm:ss");

            //            dataGridView3.Rows[0].Cells[ii + 1].Value = allSensors[j].values[(int)chart1.ChartAreas[0].CursorX.Position].value;

            //            textBox3.Text = allSensors[j].values[(int)chart1.ChartAreas[0].CursorX.Position].dt.ToString();
            //            DateTime PixelTime = new DateTime(1970, 1, 1).AddSeconds(allSensors[j].values[(int)chart1.ChartAreas[0].CursorX.Position].dt);
            //            textBox4.Text = PixelTime.ToString("HH:mm:ss");

            //            if (NumberSeriesNew > 1)
            //            {
            //                TochkaDannih.Rows.Add();
            //            }
            //        }



        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            button7.Enabled = true;
            button6.Enabled = true;
            indexPositionCursor--;
            chart1.ChartAreas[0].CursorX.Position =
                MyAllSensors[0].MyListRecordsForOneKKS[indexPositionCursor].ValueTimeForDAT;
            textBox3.Text = MyAllSensors[0].MyListRecordsForOneKKS[indexPositionCursor].DateTime.ToString();
            textBox4.Text = chart1.ChartAreas[0].CursorX.Position.ToString();
            for (int j = 0; j < MyAllSensors.Count; j++)
            {
                if (comboBox1.Text == MyAllSensors[j].KKS_Name)
                {
                    chart1.Series[19].ChartType = SeriesChartType.Point;
                    chart1.Series[19].Color = Color.Black;
                    chart1.Series[19].Points.Clear();
                    DataPoint dp = new DataPoint(chart1.ChartAreas[0].CursorX.Position,
                        MyAllSensors[j].MyListRecordsForOneKKS[indexPositionCursor].Value);
                    dp.MarkerStyle = MarkerStyle.Cross;
                    dp.MarkerSize = 10;
                    dp.IsValueShownAsLabel = true;
                    chart1.Series[19].Points.Add(dp);
                }
            }

        }

        private List<int> indexes = new List<int>();
        //    List<int> BeginIndex = new List<int>();
        //   List<int> EndIndex = new List<int>();
        private double position;
        private List<double> Time_Per = new List<double>();
        // List<double> myTok = new List<double>();
        // double beg_per;
        // double end_per;
        //   myOneVozmuchenie.N_Per=0;
        private void button5_Click(object sender, EventArgs e)
        {
            button6.Enabled = true;
            button7.Enabled = true;
            indexPositionCursor++;
            chart1.ChartAreas[0].CursorX.Position =
                MyAllSensors[0].MyListRecordsForOneKKS[indexPositionCursor].ValueTimeForDAT;
            textBox3.Text = MyAllSensors[0].MyListRecordsForOneKKS[indexPositionCursor].DateTime.ToString();
            textBox4.Text = chart1.ChartAreas[0].CursorX.Position.ToString();
            for (int j = 0; j < MyAllSensors.Count; j++)
            {
                if (comboBox1.Text == MyAllSensors[j].KKS_Name)
                {
                    chart1.Series[19].ChartType = SeriesChartType.Point;
                    chart1.Series[19].Color = Color.Black;
                    chart1.Series[19].Points.Clear();
                    DataPoint dp = new DataPoint(chart1.ChartAreas[0].CursorX.Position,
                        MyAllSensors[j].MyListRecordsForOneKKS[indexPositionCursor].Value);
                    dp.MarkerStyle = MarkerStyle.Cross;
                    dp.MarkerSize = 10;
                    dp.IsValueShownAsLabel = true;
                    chart1.Series[19].Points.Add(dp);
                }
            }
        }

        private void sddsToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private double indexPoint;
        private int index;

        private int IndexReac()
        {
            index = 0;
            for (int i = 0; i < MyAllSensors.Count; i++)
            {
                //  if (comboBox2.Text == MyAllSensors[i].KKS_Name)
                {
                    index = i;
                }
            }
            return index;
        }

        public Chart MaxAxisY()
        {
            return chart1;
        }

        private int indexPositionCursor;

        private void chart1_MouseDown_1(object sender, MouseEventArgs e)
        {
            button6.Enabled = false;
            button7.Enabled = false;
            if (e.Button == MouseButtons.Right)
            {
                button6.Enabled = true;
                chart1.ContextMenuStrip = contextMenuStrip2;
            }
            if (e.Button == MouseButtons.Left)
            {
                if (e.Y < 450 && e.Y > 0 && e.X > 0 && e.X < 117)
                {
                    if (Graph._numberseries > 0)
                    {
                        MinMaxX ChangeParametrImage = new MinMaxX();
                        ChangeParametrImage.Owner = this;
                        // ChangeParametrImage.textBox1.Text = chart1.ChartAreas[0].AxisY.Maximum.ToString();
                        // ChangeParametrImage.textBox2.Text = chart1.ChartAreas[0].AxisY.Minimum.ToString();
                        ChangeParametrImage.Show();
                        ChangeParametrImage.Location = e.Location;

                        ChangeParametrImage.textBox1.Text = chart1.ChartAreas[0].AxisY.Maximum.ToString();
                        ChangeParametrImage.textBox2.Text = chart1.ChartAreas[0].AxisY.Minimum.ToString();
                    }
                }
                //  MessageBox.Show(indexPoint.ToString());


                //for (int j = 0; j < MyAllSensors.Count; j++)
                //  {
                //  if (comboBox1.Text == MyAllSensors[j].KKS_Name)
                // {
                //    chart1.Series[19].ChartType = SeriesChartType.Point;
                //     chart1.Series[19].Color = Color.Black;
                //    chart1.Series[19].Points.Clear();
                //    position = (int)chart1.ChartAreas[0].CursorX.Position;

                //    //MessageBox.Show(((int)chart1.ChartAreas[0].CursorX.Position).ToString());
                //    textBox3.Text = MyAllSensors[j].MyListRecordsForOneKKS[position].ValueTimeForDAT.ToString();
                ////     textBox4.Text = MyAllSensors[j].MyListRecordsForOneKKS[position].DateTime.ToString("HH:mm:ss");
                //   DataPoint dp = new DataPoint(chart1.ChartAreas[0].CursorX.Position, MyAllSensors[j].MyListRecordsForOneKKS[position].Value);
                //   dp.MarkerStyle = MarkerStyle.Cross;
                //  dp.MarkerSize = 10;
                //  dp.IsValueShownAsLabel = true;
                //  chart1.Series[19].Points.Add(dp);
                //}
                // }
                //  chart1.ChartAreas[0].CursorX.Position++;
                try
                {
                    position = chart1.ChartAreas[0].CursorX.Position;
                    //   MessageBox.Show(position.ToString());
                    List<double> dTDoubles = new List<double>();
                    dTDoubles.Clear();
                    for (int i = 0; i < _timeList.Count; i++)
                    {
                        dTDoubles.Add(Math.Abs(_timeList[i] - position));
                    }
                    indexPositionCursor = dTDoubles.IndexOf(dTDoubles.Min());
                }
                catch
                {

                }

            }

        }

        private bool ter = false;
        //    int RET = 1; 
        private void button3_Click_2(object sender, EventArgs e)
        {
            //  MessageBox.Show(tableLayoutPanel10.ColumnStyles[2].Width.ToString());

            if (ter)
            {
                tableLayoutPanel10.ColumnStyles[3].Width = 0;
                tableLayoutPanel10.ColumnStyles[2].Width = 2F;
                // RET++;
                ter = false;
            }

            else
            {
                tableLayoutPanel10.ColumnStyles[3].Width = 86F;
                tableLayoutPanel10.ColumnStyles[2].Width = 3F;
                ter = true;
                double timestamp1 = chart1.ChartAreas[0].AxisX.Minimum;
                DateTime dateBeg = new DateTime(1970, 1, 1, 0, 0, 0, 0).AddSeconds(timestamp1);
                double timestamp2 = chart1.ChartAreas[0].AxisX.Maximum;
                DateTime dateEnd = new DateTime(1970, 1, 1, 0, 0, 0, 0).AddSeconds(timestamp2);
                //    textBox14.Text = dateBeg.ToString();
                //    textBox15.Text = dateEnd.ToString();
                dateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
                dateTimePicker1.CustomFormat = "dd.MM.yy HH:mm:ss";
                dateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
                dateTimePicker2.CustomFormat = "dd.MM.yy HH:mm:ss";
                dateTimePicker1.Value = DateTime.FromOADate(dateBeg.ToOADate());
                dateTimePicker2.Value = DateTime.FromOADate(dateEnd.ToOADate());
                //   MessageBox.Show(main.chart1.ChartAreas[0].AxisX.Maximum.ToString());
                //    textBox14.Text = chart1.ChartAreas[0].AxisX.Minimum.ToString();
                //    textBox15.Text = chart1.ChartAreas[0].AxisX.Maximum.ToString();
                //double timestamp = double.Parse(massiv_znacheniy_postrochno[0].Replace('.', ',').Trim());
                //DateTime date = new DateTime(1970, 1, 1, 0, 0, 0, 0).AddSeconds(timestamp);
                ////  DateTime WindowsTime = new DateTime(1970, 1, 1,1,1,1).AddSeconds(double.Parse(massiv_znacheniy_postrochno[0].Replace('.', ',').Trim()));
                //onerec.DateTime = date;
            }

            //   tableLayoutPanel10.ColumnStyles[2].Width = 100;
            //  tableLayoutPanel10.it
        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                if (double.Parse(textBox1.Text) < chart1.ChartAreas[0].AxisY.Maximum)
                {
                    chart1.ChartAreas[0].AxisY.Minimum = double.Parse(textBox1.Text);
                }

            }
            catch
            {

            }
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (double.Parse(textBox5.Text) >= chart1.ChartAreas[0].AxisY.Minimum)
                {
                    chart1.ChartAreas[0].AxisY.Maximum = double.Parse(textBox5.Text);
                }
            }
            catch
            {

            }
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (double.Parse(textBox12.Text) < chart1.ChartAreas[0].AxisY2.Maximum)
                {
                    chart1.ChartAreas[0].AxisY2.Minimum = double.Parse(textBox12.Text);
                }

            }
            catch
            {

            }
        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (double.Parse(textBox11.Text) >= chart1.ChartAreas[0].AxisY2.Minimum)
                {
                    chart1.ChartAreas[0].AxisY2.Maximum = double.Parse(textBox11.Text);
                }
            }
            catch
            {

            }
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            //  Form1 main = this.Owner as Form1;
            try
            {
                chart1.ChartAreas[0].AxisY.Interval = double.Parse(textBox6.Text);
            }
            catch
            {

            }
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            try
            {
                chart1.ChartAreas[0].AxisY2.Interval = double.Parse(textBox10.Text);
            }
            catch
            {

            }
        }

        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                chart1.ChartAreas[0].AxisY.LabelAutoFitMinFontSize = (int) numericUpDown4.Value;
            }
            catch
            {

            }
        }

        private void numericUpDown6_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                chart1.ChartAreas[0].AxisY2.LabelAutoFitMinFontSize = (int) numericUpDown6.Value;
            }
            catch
            {

            }
        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {

            try
            {
                chart1.Titles.Clear();
                chart1.Titles.Add(textBox13.Text);
            }
            catch
            {

            }
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            try
            {
                chart1.ChartAreas[0].AxisX.Title = textBox9.Text;
                //       main.chart1.ChartAreas[0].AxisX.TitleFont = new Font("Times New Roman", (int)numericUpDown2.Value, FontStyle.Regular);

            }
            catch
            {

            }
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

            try
            {
                //   main.chart1.ChartAreas[0].AxisY.Title = textBox8.Text;
                chart1.ChartAreas[0].AxisY.Title = textBox8.Text;
                //       main.chart1.ChartAreas[0].AxisX.TitleFont = new Font("Times New Roman", (int)numericUpDown2.Value, FontStyle.Regular);

            }
            catch
            {

            }
        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //   main.chart1.ChartAreas[0].AxisY.Title = textBox8.Text;
                //   chart1.ChartAreas[0].AxisY2.Title = textBox14.Text;
                //       main.chart1.ChartAreas[0].AxisX.TitleFont = new Font("Times New Roman", (int)numericUpDown2.Value, FontStyle.Regular);

            }
            catch
            {

            }
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                chart1.Titles[0].Font = new System.Drawing.Font("Times New Roman", (int) numericUpDown1.Value,
                    FontStyle.Regular);
            }
            catch
            {

            }
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {

            try
            {
                chart1.ChartAreas[0].AxisX.TitleFont = new System.Drawing.Font("Times New Roman",
                    (int) numericUpDown2.Value, FontStyle.Regular);
            }
            catch
            {

            }
        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {

            try
            {
                chart1.ChartAreas[0].AxisY.TitleFont = new System.Drawing.Font("Times New Roman",
                    (int) numericUpDown3.Value, FontStyle.Regular);
            }
            catch
            {

            }
        }

        private void numericUpDown7_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                // chart1.ChartAreas[0].AxisY2.TitleFont = new System.Drawing.Font("Times New Roman", (int)numericUpDown7.Value, FontStyle.Regular);
            }
            catch
            {

            }
        }

        private void checkBox3_CheckedChanged_1(object sender, EventArgs e)
        {
            try
            {

                if (checkBox3.Checked == true)
                {
                    chart1.Legends[0].Enabled = true;
                }
                if (checkBox3.Checked == false)
                {
                    chart1.Legends[0].Enabled = false;
                }
            }
            catch
            {

            }
        }

        private void checkBox4_CheckedChanged_2(object sender, EventArgs e)
        {

            if (checkBox4.Checked == true)
            {
                chart1.ChartAreas[0].AxisX.MinorTickMark.Enabled = true;
                chart1.ChartAreas[0].AxisX.MinorTickMark.LineDashStyle = ChartDashStyle.Dash;

                chart1.ChartAreas[0].AxisX.MinorGrid.Enabled = true;
                chart1.ChartAreas[0].AxisX.MajorGrid.Enabled = true;
                chart1.ChartAreas[0].AxisX.MinorGrid.LineDashStyle = ChartDashStyle.Dash;
                chart1.ChartAreas[0].AxisX.MinorGrid.LineColor = Color.Gray;
                chart1.ChartAreas[0].AxisX.MinorTickMark.LineColor = Color.Gray;
            }
            if (checkBox4.Checked == false)
            {
                chart1.ChartAreas[0].AxisX.MinorTickMark.Enabled = false;
                chart1.ChartAreas[0].AxisX.MinorGrid.Enabled = false;
            }
        }

        private void checkBox2_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                chart1.ChartAreas[0].AxisY.MinorTickMark.Enabled = true;
                chart1.ChartAreas[0].AxisY.MinorTickMark.LineDashStyle = ChartDashStyle.Dash;

                chart1.ChartAreas[0].AxisY.MinorGrid.Enabled = true;
                chart1.ChartAreas[0].AxisY.MinorGrid.LineDashStyle = ChartDashStyle.Dash;
                chart1.ChartAreas[0].AxisY.MinorGrid.LineColor = Color.Gray;
                chart1.ChartAreas[0].AxisY.MinorTickMark.LineColor = Color.Gray;
            }
            if (checkBox2.Checked == false)
            {
                chart1.ChartAreas[0].AxisY.MinorTickMark.Enabled = false;
                chart1.ChartAreas[0].AxisY.MinorGrid.Enabled = false;
            }
        }

        private void radioButton1_CheckedChanged_1(object sender, EventArgs e)
        {

            chart1.ChartAreas[0].AxisX.MinorTickMark.LineDashStyle = ChartDashStyle.Dash;
            chart1.ChartAreas[0].AxisX.MinorGrid.LineDashStyle = ChartDashStyle.Dash;
            chart1.ChartAreas[0].AxisX.MinorGrid.LineColor = Color.LightGray;
            chart1.ChartAreas[0].AxisX.MinorTickMark.LineColor = Color.LightGray;
        }

        private void radioButton2_CheckedChanged_1(object sender, EventArgs e)
        {

            chart1.ChartAreas[0].AxisX.MinorTickMark.LineDashStyle = ChartDashStyle.Dot;
            chart1.ChartAreas[0].AxisX.MinorGrid.LineDashStyle = ChartDashStyle.Dot;
            chart1.ChartAreas[0].AxisX.MinorGrid.LineColor = Color.LightGray;
            chart1.ChartAreas[0].AxisX.MinorTickMark.LineColor = Color.LightGray;
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {

            chart1.ChartAreas[0].AxisY.MinorTickMark.LineDashStyle = ChartDashStyle.Dash;
            chart1.ChartAreas[0].AxisY.MinorGrid.LineDashStyle = ChartDashStyle.Dash;
            chart1.ChartAreas[0].AxisY.MinorGrid.LineColor = Color.LightGray;
            chart1.ChartAreas[0].AxisY.MinorTickMark.LineColor = Color.LightGray;
        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        private List<int> ForERORR = new List<int>();

        private List<int> indexPositionCursorList = new List<int>();

        private void button6_Click(object sender, EventArgs e)
        {
            indexPositionCursorList.Add(indexPositionCursor);
            //  MessageBox.Show(indexPositionCursorList.Count.ToString() + " " + indexPositionCursor);



            // ForERORR.Add(0);
            // if (BeginIndex.Count == 0)
            // {
            //   pdH.Add(0);
            //}




            //   Time_Per.Add(allSensors[0].MyListRecordsForOneKKS[(int)chart1.ChartAreas[0].CursorX.Position].value1);

            //   MessageBox.Show("gv");



            //      BeginIndex.Add(position);

            //   if (BeginIndex.Count == 0)
            // {
            //    pdH.Add(0);
            //}
            //  if (BeginIndex.Count > 2)
            //   {
            //     pdH.Add(step);
            //   }
            //  pdH.Add(step);

            //   if (BeginIndex.Count>1)
            {
                //      No_Perem_H();
                //   Time_Per.Clear();
                //    }


                button6.Visible = false;
                button7.Visible = true;
                //   button6.Enabled = false;
                button7.Enabled = false;
                button7.BackColor = Color.Yellow;
                // No_Perem_H();
                //  MessageBox.Show(BeginIndex[0].ToString());
                // myOneVozmuchenie.N_Per = 0;
                //       MessageBox.Show(MyTok.Count.ToString() + " " + myReactivity.Count.ToString() + " " + pdH.Count.ToString());
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            indexPositionCursorList.Add(indexPositionCursor);
            // MessageBox.Show(indexPositionCursorList.Count.ToString() + " " + indexPositionCursor);
            //   EndIndex.Add(position);
            //   ForERORR.Add(1);
            //  if(EndIndex.Count == 1)
            // {
            //  button8.Enabled = true; //   Time_Per.RemoveAt(1);
            // }
            // Razdroblenie_H();
            // Time_Per.Clear();

            button6.Visible = true;
            button7.Visible = false;
            button6.Enabled = false;
            button6.BackColor = Color.Yellow;
            button8.Enabled = true;
            //     MessageBox.Show(EndIndex[EndIndex.Count-1].ToString());
            //  MessageBox.Show(myOneVozmuchenie.N_Per.ToString());
            //  MessageBox.Show(Time_Per.Count.ToString());
        }

        //  List<double> MyTok = new List<double>();
        // List<double> MyTime = new List<double>();
        private List<double> proba = new List<double>();

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        // int NOper = -2;
        // private OneVozmuchenieData myOneVozmuchenie = new OneVozmuchenieData();
        // List<double> pdH = new List<double>();
        // List<double> myReactivity = new List<double>();
        // double step;
      //////  if (comboBox2.Text == MyAllSensors[i].KKS_Name)
             //////   {
                 ////   for (int j = indexPositionCursorList[0];
             //      j < indexPositionCursorList[indexPositionCursorList.Count - 1] + 1;
                //        j++)
                   /// {
                    /////    _jList.Add(MyAllSensors[i].MyListRecordsForOneKKS[j].Value/
                    ////               MyAllSensors[i].MyListRecordsForOneKKS[indexPositionCursorList[0]].Value);
                    /////    _tList.Add(MyAllSensors[i].MyListRecordsForOneKKS[j].ValueTimeForDAT);
                  /////  }

        double DroAver = 0;
        private double PoPichkam()
                   {
                   //    double r = 0;
          //  List<double> n = new List<double>();
           // double DroAver = 0;
            for (int i = 0; i < MyAllSensors.Count; i++)
            {
                if (comboBox1.Text == MyAllSensors[i].KKS_Name)
                {
                    for (int j = 1;
                   j < indexPositionCursorList.Count;
                        j++)
                    {
                        if (j % 2 != 0)
                        {


                       //     MessageBox.Show(
                          //  (MyAllSensors[i].MyListRecordsForOneKKS[indexPositionCursorList[j]].Value + " " +
                          //   MyAllSensors[i].MyListRecordsForOneKKS[indexPositionCursorList[j - 1]].Value).ToString());


                            DroAver = DroAver +
                                      (MyAllSensors[i].MyListRecordsForOneKKS[indexPositionCursorList[j]].Value -
                                       MyAllSensors[i].MyListRecordsForOneKKS[indexPositionCursorList[j - 1]].Value);



                            //ведь каждый шаг должен в сумме давать 2
                            //for (int j = myOneVozmuchenie.PerIndex[i, 0] + 1; j <= myOneVozmuchenie.PerIndex[i, 1]; j++)
                            //{
                            //    ddH -= (2 / (myOneVozmuchenie.Per[i, 1] - myOneVozmuchenie.Per[i, 0])) * (myInputData.Data[0][j] - myInputData.Data[0][j - 1]);
                            //    pdH.Add(ddH);
                            //}

                        }
                    }
                //    MessageBox.Show(DroAver.ToString());

                  //  MessageBox.Show((indexPositionCursorList.Count/2).ToString());
                    DroAver = DroAver / (indexPositionCursorList.Count/2);
                 //   MessageBox.Show(DroAver.ToString());

                    DroAver = DroAver/-2;
                 //   MessageBox.Show(DroAver.ToString());

                    //TODO: ПЕРЕВОД В ПРОЦЕНТЫ ЗАЧЕМ?
                    DroAver = DroAver * MyConst._beff;
                  //  MessageBox.Show(DroAver.ToString());

                    //double sum_dR = 0;
                    //double dRdH = 0;
                    //for (int i = 0; i < allSensors.Count; i++)
                    //{
                    //  if (comboBox2.Text == allSensors[i].KKS_Name)
                    //{
                    //  for (int j = 0; j < BeginIndex.Count; j++)
                    //{
                    //  
                    //sum_dR = sum_dR + (allSensors[i].MyListRecordsForOneKKS[EndIndex[j]].Value - allSensors[i].MyListRecordsForOneKKS[BeginIndex[j]].Value);
                    // }
                    // dRdH = sum_dR / (2 * (BeginIndex.Count));
                    //}
                    // }
                    //dRdH = sum_dR / (2 * (BeginIndex.Count));
                    //   return dRdH;
                }
            }
            return DroAver;
        }
        double DEL = 1;
        static pertubResult tempR;
        int indexSS = 1000;
        StreamWriter lll = new StreamWriter("D:\\дЛЯ ВЫВОДа.txt");
        StreamWriter lllВ = new StreamWriter("D:\\дЛЯ ПРОВЕРКИ ТЕНДЕНЦИИ.txt");
        
        List<double> SsOLD = new List<double>();
        List<double> SsNEW = new List<double>();
        List<double> TAUlist = new List<double>();
        List<double> SsList = new List<double>();
        private void SearchDiffEffect()
        {
            tempR = new pertubResult();
           // pertubResult tempR = new pertubResult();
            if (comboBox6.Text == "Подбор")
            {
                double Ss = 1000;
           //     pertubResult tempR = new pertubResult();
             //   tempR = new pertubResult();
                for (int i = 0; i < 400; i++)
                {
                    tempR = Calc(3 + i / 200.0, _tList, _jKorList, _rCalcList, _dHList);
                  //  lllВ.WriteLine(tempR.SS + " " + Ss);
                    SsOLD.Add(Ss);
                    SsNEW.Add(tempR.SS);
                    TAUlist.Add(tempR.tau);
                //    if (tempR.SS > Ss)
                   // {
                      //  MessageBox.Show(tempR.SS + " " + Ss);
                       // tempR = Calc(3 + (i - 1) / 200.0, _tList, _jKorList, _rCalcList, _dHList);
                        //TODO: 400 РАЗ ОБРАБАТЫВАЕТСЯ ОДНО И ТОЖЕ ВОЗМУЩЕНИЕ С РАЗНЫМ ПОДБОРОМ ТАУ. В МЕТОДЕ CALC ВСЕГДА НАБИРАЕТСЯ СУММА КВАДРАТОВ РАЗНОСТЕЙ РЕАКТВИНОСТЕЙ НАЗВАЕМОЙ НЕВЯЗКОЙ. 
                        //TODO: TEMPR.SS В НЕМ ЖЕ ПРИРАВНИВАЕТСЯ ЭТОЙ СУММЕ. В ИТОГЕ ПО УСЛОВИЮ И ПО ЛОГИКЕ НЕВЯЗКА ИЛИ ЖЕ СВОЙСТВО TEMPR.SS СРАВНИВАЕТСЯ C ЕГО ЖЕ ПРЕДЫДУЩЕМ ЗНАЧЕНИЕМ
                        //TODO: И КАК ТОЛЬКО ЭТО СВОЙСТВО СТАНОВИТСЯ БОЛЬШЕ ПРЕДЫДУЩЕГО ТО ПРОИСХОДИТ ОТКАТ К ПРЕДЫДУЩЕМУ ЗНАЧЕНИЮ ТАУ И АЛГОРИТМ ПРЕРЫВАЕТСЯ
                        //TODO: ВОПРОС: МОЖЕТ ЛИ БЫТЬ ПОСЛЕ СРАБАТЫВАНИЯ ВЕЛИЧИНА НЕВЯЗКИ МЕНЬШЕ ТОГО ЧТО БЫЛО ЕСЛИ АЛГОРИТМ НЕ ПРЕРВЕТСЯ ВЕДЬ ИЗ ФАЙЛА ЗАПИСИ ВИДНО ЧТО 
                    //    break;
                  //  }
                  //  lll.WriteLine(tempR.aH + " " + tempR.tau + " " + tempR.SS);
                    Ss = tempR.SS;
                }


        
             //   double DEL = 1;
                for (int i = 0; i < SsOLD.Count; i++)
                {
                    lll.WriteLine(SsNEW[i] + " " + SsOLD[i] + " " + TAUlist[i]);
                }

            //    lllВ.WriteLine(tempR.SS + " " + Ss);
                for (int i = 0; i < SsOLD.Count; i++)
                {
                    if ((SsNEW[i] - SsOLD[i])>0)
                    {
                        if (SsNEW[i] - SsOLD[i]< DEL)
                        {
                            DEL = SsNEW[i] - SsOLD[i];
                            indexSS = i;
                        }
                    }
                    SsList.Add(SsNEW[i]-SsOLD[i]);
                }



                tempR = Calc(TAUlist[indexSS], _tList, _jKorList, _rCalcList, _dHList);

                lllВ.WriteLine(SsNEW[indexSS] + " " + SsOLD[indexSS] + " " + TAUlist[indexSS]);


            }
            lll.Close();
            lllВ.Close();
            if (comboBox6.Text == "3")
            {
              //  tempR = new pertubResult();
                tempR = Calc(3, _tList, _jKorList, _rCalcList, _dHList);
            }
            if (comboBox6.Text == "4")
            {
              //  tempR = new pertubResult();
                tempR = Calc(4, _tList, _jKorList, _rCalcList, _dHList);
            }
            if (comboBox6.Text == "5")
            {
              //  tempR = new pertubResult();
                tempR = Calc(5, _tList, _jKorList, _rCalcList, _dHList);
            }
            tempR.Ro = tempR.Ro * MyConst._beff;
            tempR.aH = tempR.aH * MyConst._beff;
            tempR.b = tempR.b*MyConst._beff;
            tempR.SS = tempR.SS * MyConst._beff;
        }

        private void button8_Click(object sender, EventArgs e)
        {


            //   tempR.Ro = tempR.Ro * MyConst.Rect.Beff;
            //   tempR.aH = tempR.aH * MyConst.Rect.Beff;
            //    tempR.b = tempR.b * MyConst.Rect.Beff;
            //  tempR.SS = tempR.SS * MyConst.Rect.Beff;
        }

        private void button4_Click_1(object sender, EventArgs e)
        {

        }

        private void button9_Click(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel7_Paint(object sender, PaintEventArgs e)
        {

        }

        private void checkBox1_CheckedChanged_1(object sender, EventArgs e)
        {

        }

        private void button2_Click_2(object sender, EventArgs e)
        {
            MyListOfSensors N = new MyListOfSensors();
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                foreach (var item in openFileDialog1.FileNames)
                {
                    N.LoadFromFile(item, MyAllSensors);
                }
            }
            if (MyAllSensors.Count > 0)
            {
                добавитьНаОсьXToolStripMenuItem.Enabled = true;
            }
            dataGridView4.Columns.Clear();
            dataGridView4.Columns.Add("Время", "Время");
            dataGridView4.Columns.Add("Ток(эксперимент), А", "Ток(эксперимент), А");
            dataGridView4.Columns.Add("\u03C1, %", "\u03C1, %");
            dataGridView4.Columns.Add("N, МВт", "N, МВт");
            dataGridView4.Columns.Add("H12, см", "H12, см");
            dataGridView4.Columns.Add("T, \u2103", "T, \u2103");

            dataGridView3.Columns.Clear();
            dataGridView3.Columns.Add("Время", "Время");
            dataGridView3.Columns.Add("Ток(эксперимент), А", "Ток(эксперимент), А");
            dataGridView3.Columns.Add("\u03C1, %", "\u03C1, %");
            dataGridView3.Columns.Add("N, МВт", "N, МВт");
            dataGridView3.Columns.Add("H12, см", "H12, см");
            dataGridView3.Columns.Add("T, \u2103", "T, \u2103");
            dataGridView3.Columns.Add("Ток(относительный)", "Ток(относительный)");
            dataGridView3.Columns.Add("ПЭ", "ПЭ");


            try
            {
                dataGridView4.Visible = false;

                this.dataGridView4.CellValueNeeded += new
                    DataGridViewCellValueEventHandler(dataGridView4_CellValueNeeded);

                dataGridView4.Rows.Clear();
                customers1.Clear();

                for (int i = 0; i < MyAllSensors[0].MyListRecordsForOneKKS.Count; i++)
                {
                    this.customers1.Add(
                        new MyVirtualClass(MyAllSensors[0].MyListRecordsForOneKKS[i].DateTime.ToString("HH:mm:ss.fff")));
                }

                if (this.dataGridView4.RowCount == 0)
                {
                    this.dataGridView4.RowCount = 1;
                }

                this.dataGridView4.RowCount = MyAllSensors[0].MyListRecordsForOneKKS.Count;
                dataGridView4.Visible = true;
            }
            catch
            {
                //  MessageBox.Show(ex0.Message);
            }

            dataGridView4.Columns[indexData4].Width = 100;

            checkedListBox1.Items.Clear();
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();
            comboBox3.Items.Clear();
            comboBox4.Items.Clear();
            comboBox5.Items.Clear();
            foreach (Sensors item in MyAllSensors)
            {
                checkedListBox1.Items.Add(item.KKS_Name);
                comboBox1.Items.Add(item.KKS_Name);
                comboBox2.Items.Add(item.KKS_Name);
                comboBox3.Items.Add(item.KKS_Name);
                comboBox4.Items.Add(item.KKS_Name);
                comboBox5.Items.Add(item.KKS_Name);
            }
        }

        private bool flag = true;

        private void button9_Click_1(object sender, EventArgs e)
        {
            //  button9.BackColor = Color.Yellow;
            if (flag)
            {
                // button9.BackColor = Color.Yellow;
                chart1.ChartAreas[0].CursorX.IsUserEnabled = true;
                chart1.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
                chart1.ChartAreas[0].AxisX.ScaleView.Zoomable = true;
                chart1.ChartAreas[0].AxisX.ScrollBar.IsPositionedInside = true;
                chart1.ChartAreas[0].CursorY.IsUserEnabled = true;
                chart1.ChartAreas[0].CursorY.LineColor = Color.Blue;
                chart1.ChartAreas[0].CursorX.LineColor = Color.Blue;
                chart1.ChartAreas[0].CursorY.LineWidth = 2;
                chart1.ChartAreas[0].CursorX.LineWidth = 2;
                chart1.ChartAreas[0].CursorY.LineDashStyle = ChartDashStyle.Dash;
                chart1.ChartAreas[0].CursorX.LineDashStyle = ChartDashStyle.Dash;

                //   chart1.ChartAreas[0].CursorY.IsUserEnabled = true;

                //  chart1.ChartAreas[0].AxisX.ScaleView.Zoomable = true;
                chart1.ChartAreas[0].AxisY.ScaleView.Zoomable = true;
                chart1.ChartAreas[0].CursorY.Interval = 0.0000001;
                chart1.ChartAreas[0].CursorX.Interval = 0.0000001;
                // Set automatic scrochart1.ChartAreas[0].AxisYlling 
                //  chart1.ChartAreas[0].CursorX.AutoScroll = true;
                chart1.ChartAreas[0].CursorY.AutoScroll = true;

                // Allow user selection for Zoom
                //    chart1.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
                chart1.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;


                //    chart1.ChartAreas[0].CursorX.SelectionColor = Color.Blue;
                chart1.ChartAreas[0].CursorY.SelectionColor = Color.Blue;
                flag = false;
            }
            else
            {
                chart1.ChartAreas[0].CursorX.IsUserSelectionEnabled = false;
                chart1.ChartAreas[0].CursorY.IsUserSelectionEnabled = false;
                chart1.ChartAreas[0].CursorY.AutoScroll = false;
                chart1.ChartAreas[0].AxisY.ScaleView.Zoomable = false;
                flag = true;
            }
        }

        private void button4_Click_2(object sender, EventArgs e)
        {
            добавитьНаОсьXToolStripMenuItem.Enabled = false;
            добавитьНаДополнительнуюОсьYОтВремениToolStripMenuItem.Enabled = false;

            Graph.MinmaxListPrimary.Clear();
            Graph.MinmaxListSecondary.Clear();

            //   NumberSeries = 0;
            Graph._numberseries = 0;
            chart1.Titles.Clear();
            //     добавитьНаОсьXToolStripMenuItem.Enabled = false;

            //   checkBox4.Enabled = true;
            //  открытьToolStripMenuItem.Enabled = true;
            // checkedListBox1.Items.Clear();
            checkedListBox1.Items.Clear();
            //   MyAllSensors.Clear();
            MyAllSensors.Clear();
            //  chart1.Legends.Clear();
            //   checkedListBox1.Items.Clear();
            //   chart1.Series.Clear();
            chart1.ChartAreas[0].Position.Auto = true;
            dataGridView4.Rows.Clear();
            dataGridView3.Rows.Clear();


            for (int i = 0; i < chart1.Series.Count; i++)
            {
                chart1.Series[i].IsVisibleInLegend = false;
                chart1.Series[i].Points.Clear();
            }


        }

        private void добавитьНаДополнительнуюОсьYОтВремениToolStripMenuItem_Click(object sender, EventArgs e)
        {
            flagAxis = false;
            Graph.CreateLine(MyAllSensors, checkedListBox1.Text, chart1, flagAxis);
            button6.Enabled = true;
        }

        private void chart1_MouseMove(object sender, MouseEventArgs e)
        {
            try
            {
                if (Graph._numberseries > 0)
                {
                    textBox2.Text = chart1.ChartAreas[0].AxisX.PixelPositionToValue(e.X).ToString();
                    if (checkBox1.Checked == false)
                    {
                        textBox14.Text = chart1.ChartAreas[0].AxisY.PixelPositionToValue(e.Y).ToString();
                    }
                    if (checkBox1.Checked == true)
                    {
                        textBox14.Text = chart1.ChartAreas[0].AxisY2.PixelPositionToValue(e.Y).ToString();
                    }

                    // if (e.Button == MouseButtons.Left) //проверка нажатия левой кнопки
                    // {
                    // label26.Text = string.Format("x= " + e.X);
                    //  label27.Text = string.Format("y= " + e.Y);


                    //  }
                }
            }
            catch (Exception)
            {

            }
        }

        private int indexData4 = 0;

        private string comboBoxText = null;

        public static List<double> _PEList = new List<double>();
        public static List<double> _jExpList = new List<double>();
        public static List<double> _jKorList = new List<double>();
        public static List<double> _jRealList = new List<double>();
        public static List<double> _rExpList = new List<double>();
        public static List<double> _rCalcList = new List<double>();
        public static List<double> _dHList = new List<double>();
        public static List<double> _HList = new List<double>();
        public static List<double> _tList = new List<double>();

        double PE;
        
        private void button8_Click_2(object sender, EventArgs e)
        {

            //    indexPE = 0;
            // MessageBox.Show(_PEList[indexPE].ToString());
            // indexPE++;
            //  MessageBox.Show(_PEList[indexPE].ToString());

            //TODO: РАСЧЕТ DDH
            double ddH = 0;
            _dHList.Add(ddH);
            int step = 0;

            for (int j = 1; j < indexPositionCursorList.Count; j++)
            {
                if (j%2 != 0)
                {
                    for (int i = indexPositionCursorList[j - 1];
                        i < indexPositionCursorList[j];
                        i++)
                    {
                        ddH -= (2/
                                (MyAllSensors[0].MyListRecordsForOneKKS[indexPositionCursorList[j]].ValueTimeForDAT -
                                 MyAllSensors[0].MyListRecordsForOneKKS[indexPositionCursorList[j - 1]].ValueTimeForDAT))*
                               (MyAllSensors[0].MyListRecordsForOneKKS[i + 1].ValueTimeForDAT -
                                MyAllSensors[0].MyListRecordsForOneKKS[i].ValueTimeForDAT);
                        _dHList.Add(ddH);
                    }
                }
                //        MessageBox.Show(ddH.ToString());
                if (j%2 == 0)
                {
                    for (int i = indexPositionCursorList[j - 1];
                        i < indexPositionCursorList[j];
                        i++)
                    {
                        _dHList.Add(ddH);
                    }
                }
            }

            for (int i = 0; i < 200; i++)
            {
                _dHList.Add(ddH);
            }


            int indexDDH = 0;
           
            //TODO: ДОБАВЛЕНИЕ ЗНАЧЕНИЙ ТОКА И ВРЕМЕНИ ПО ОТМЕЧЕННЫМ ТОЧКАМ
            for (int i = 0; i < MyAllSensors.Count; i++)
            {
                if (comboBox2.Text == MyAllSensors[i].KKS_Name)
                {
                    for (int j = indexPositionCursorList[0];
                        j < indexPositionCursorList[indexPositionCursorList.Count - 1] + 201;
                        j++)
                    {
                        double Inext = MyAllSensors[i].MyListRecordsForOneKKS[j].Value/
                                       MyAllSensors[i].MyListRecordsForOneKKS[indexPositionCursorList[0]].Value;
                        //   _jExpList.Add(Inext);

                        //  MessageBox.Show(_dHList[indexDDH] + " " + _dHList[indexDDH - 1]);
                        if (_jExpList.Count > 0 && indexDDH > 0)
                        {
                            _jKorList.Add(Inext - (PE*(_dHList[indexDDH])));
                        }
                        if (_jExpList.Count == 0)
                        {
                            _jKorList.Add(Inext);
                        }
                        //   _jKorList.Add(Inext - (PE * (_dHList[indexDDH] - _dHList[indexDDH-1])));
                        //     MessageBox.Show(indexDDH.ToString());
                        _jExpList.Add(Inext);
                        _jRealList.Add(MyAllSensors[i].MyListRecordsForOneKKS[j].Value);
                        //  MessageBox.Show(_jExpList.Count.ToString());
                        _tList.Add(MyAllSensors[i].MyListRecordsForOneKKS[j].ValueTimeForDAT);
                        indexDDH++;
                        // MessageBox.Show(indexDDH.ToString());
                    }
                }

                if (comboBox1.Text == MyAllSensors[i].KKS_Name)
                {
                    for (int j = indexPositionCursorList[0];
                        j < indexPositionCursorList[indexPositionCursorList.Count - 1] + 201;
                        j++)
                    {
                        _rExpList.Add(MyAllSensors[i].MyListRecordsForOneKKS[j].Value);
                        // _rExpList.Add(0);
                    }
                }


                //  if (comboBox5.Text == MyAllSensors[i].KKS_Name)
                //{
                //for (int j = indexPositionCursorList[0];
                //    j < indexPositionCursorList[indexPositionCursorList.Count - 1] + 1;
                //    j++)
                // {
                //     _HList.Add(MyAllSensors[i].MyListRecordsForOneKKS[j].Value);
                // }
                // }
            }

            Current Y = new Current();
            Y.AddData(_jKorList, _rExpList, _tList, _rCalcList);


            //TODO: ЭТО УЖЕ РАСЧЕТ ДИФФ-ЭФФЕКТА В КОНЦЕ В САМОМ 
            //    pertubResult t = new pertubResult();
            SearchDiffEffect();


            Series RCalc = new Series();
            RCalc.Name = "Rcorr" + Graph._numberseries;
            RCalc.ChartType = SeriesChartType.Line;
            RCalc.Points.Clear();
            for (int i = 0; i < _tList.Count; i++)
            {
                RCalc.Points.AddXY(
                    _tList[i],
                    _rCalcList[i]);
            }

            
            RCalc.BorderWidth = 3;
            chart1.Series.Add(RCalc);
            Graph._numberseries++;


            Series dH = new Series();
            dH.Name = "dH" + Graph._numberseries;
            dH.ChartType = SeriesChartType.Line;
            dH.Points.Clear();
            for (int i = 0; i < _tList.Count; i++)
            {
                dH.Points.AddXY(
                    _tList[i],
                    _dHList[i]);
            }

            dH.YAxisType = AxisType.Secondary;
            dH.BorderWidth = 3;
      //      dH.Color = Color.Black;
            chart1.Series.Add(dH);
            Graph._numberseries++;


            //     dataGridView2.Rows.Add(t.aH, dRdH());
            //   dataGridView2.Rows[0].DefaultCellStyle.BackColor = Color.Yellow;

            // MessageBox.Show(_jExpList.Count.ToString() + " " + _jKorList.Count + " " + _tList.Count + " " + _rExpList.Count + " " + tempR.FF.Count + " " + _dHList.Count + " " + _rCalcList.Count);
            for (int i = 0; i < _jExpList.Count; i++)
            {
                dataGridView5.Rows.Add(_tList[i], _jExpList[i], _jKorList[i], _rExpList[i], _rCalcList[i], _dHList[i],
                    tempR.FF[i]);
                dataGridView5.Rows[i].DefaultCellStyle.BackColor = Color.Turquoise;
            }



            //   MessageBox.Show(tempR.FF.Count.ToString() + " " + _tList.Count.ToString());
            //    StreamWriter WriterAH = new StreamWriter();
            dataGridView2.Rows.Add(tempR.aH, PoPichkam(), tempR.tau);

            StreamWriter b = new StreamWriter("D:\\1.txt");
            for (int i = 0; i < _tList.Count; i++)
            {
                b.WriteLine(_tList[i] + "\t" + _rCalcList[i] + "\t" + _jExpList[i] * _jRealList[0]);
            }
            b.Close();

            indexPositionCursorList.Clear();
            _jKorList.Clear();
            _tList.Clear();
            tempR.FF.Clear();
            _rExpList.Clear();
            _jExpList.Clear();
            _rCalcList.Clear();
            _dHList.Clear();
            button8.Enabled = false;
            //  dataGridView5.Rows.Clear();
            //  dataGridView2.Rows.Clear();
        }

        private int _indexRow = 0;

        private void button10_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(_indexRow.ToString());
            List<double> myValue = new List<double>();
            myValue.Clear();
            myValue.Add(MyAllSensors[0].MyListRecordsForOneKKS[indexPositionCursor].ValueTimeForDAT);
            for (int i = 0; i < MyAllSensors.Count; i++)
            {
                if (comboBox2.Text == MyAllSensors[i].KKS_Name)
                {
                    myValue.Add(MyAllSensors[i].MyListRecordsForOneKKS[indexPositionCursor].Value);
                    break;
                }
            }
            for (int i = 0; i < MyAllSensors.Count; i++)
            {
                if (comboBox1.Text == MyAllSensors[i].KKS_Name)
                {
                    myValue.Add(MyAllSensors[i].MyListRecordsForOneKKS[indexPositionCursor].Value);
                    break;
                }
            }
            for (int i = 0; i < MyAllSensors.Count; i++)
            {
                if (comboBox3.Text == MyAllSensors[i].KKS_Name)
                {
                    myValue.Add(MyAllSensors[i].MyListRecordsForOneKKS[indexPositionCursor].Value);
                    break;
                }
            }
            for (int i = 0; i < MyAllSensors.Count; i++)
            {
                if (comboBox5.Text == MyAllSensors[i].KKS_Name)
                {
                    myValue.Add(MyAllSensors[i].MyListRecordsForOneKKS[indexPositionCursor].Value);
                    break;
                }
            }
            for (int i = 0; i < MyAllSensors.Count; i++)
            {
                if (comboBox4.Text == MyAllSensors[i].KKS_Name)
                {
                    myValue.Add(MyAllSensors[i].MyListRecordsForOneKKS[indexPositionCursor].Value);
                    break;
                }
            }
            dataGridView3.Rows.Add();
            for (int i = 0; i < myValue.Count; i++)
            {
                dataGridView3.Rows[_indexRow].Cells[i].Value = myValue[i];

            }
            _indexRow++;
        }

        private void button11_Click(object sender, EventArgs e)
        {
            List<double> MyIlist = new List<double>();
            MyIlist.Clear();
            for (int i = 0; i < dataGridView3.Rows.Count - 1; i++)
            {
                if (checkBox6.Checked == false)
                {
                    MyIlist.Add((double.Parse(dataGridView3.Rows[i].Cells[1].Value.ToString().Trim())/
                                 double.Parse(dataGridView3.Rows[0].Cells[1].Value.ToString().Trim())) *
                                (double.Parse(dataGridView3.Rows[i].Cells[3].Value.ToString().Trim())/
                                 double.Parse(dataGridView3.Rows[0].Cells[3].Value.ToString().Trim())) *
                                (1 +
                                 0.01*
                                 (double.Parse(dataGridView3.Rows[0].Cells[5].Value.ToString().Trim()) -
                                  double.Parse(dataGridView3.Rows[i].Cells[5].Value.ToString().Trim()))));
                }
                if (checkBox6.Checked == true)
                {
                    MyIlist.Add((double.Parse(dataGridView3.Rows[i].Cells[1].Value.ToString().Trim()) /
                                 double.Parse(dataGridView3.Rows[0].Cells[1].Value.ToString().Trim()))*
                                ((double.Parse(dataGridView3.Rows[i].Cells[3].Value.ToString().Trim())*32)/
                                 (double.Parse(dataGridView3.Rows[0].Cells[3].Value.ToString().Trim())*32))*
                                (1 +
                                 0.01*
                                 (double.Parse(dataGridView3.Rows[0].Cells[5].Value.ToString().Trim()) -
                                  double.Parse(dataGridView3.Rows[i].Cells[5].Value.ToString().Trim()))));
                }
            }
            for (int i = 0; i < MyIlist.Count; i++)
            {
                dataGridView3.Rows[i].Cells[6].Value = MyIlist[i];
            }
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {


                if (Graph._numberseries > 0)
                {
                    // if (comboBox6.SelectedIndex == 0)
                    {
                        chart1.ChartAreas[0].AxisX.Minimum = chart1.Series[0].Points[0].XValue;
                        chart1.ChartAreas[0].AxisX.Maximum =
                            chart1.Series[0].Points[chart1.Series[0].Points.Count - 1].XValue;
                    }
                    // else
                    // {
                    //   FillChart();
                    //     chart1.ChartAreas[0].AxisX.Maximum = chart1.ChartAreas[0].AxisX.Minimum +
                    //   MyConst.XintervalVal[comboBox6.SelectedIndex];
                    //  chart1.ChartAreas[0].AxisX.ScrollBar.Enabled = true;
                    // chart1.ChartAreas[0].AxisX.ScrollBar.IsPositionedInside = true;
                    //  }

                }
            }
            catch
            {

            }
        }

        private void hScrollBar1_Scroll(object sender, ScrollEventArgs e)
        {
            if (Graph._numberseries > 0)
            {
                if (e.NewValue > e.OldValue)
                {
                    chart1.ChartAreas[0].AxisX.Minimum = chart1.ChartAreas[0].AxisX.Minimum + 1;
                    chart1.ChartAreas[0].AxisX.Maximum = chart1.ChartAreas[0].AxisX.Maximum + 1;
                }
                if (e.NewValue < e.OldValue)
                {
                    chart1.ChartAreas[0].AxisX.Minimum = chart1.ChartAreas[0].AxisX.Minimum - 1;
                    chart1.ChartAreas[0].AxisX.Maximum = chart1.ChartAreas[0].AxisX.Maximum - 1;
                }

            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
          //  List<double> MyRlist = new List<double>();
            _PEList.Clear();
           // MyRlist.Clear();
            for (int i = 0; i < dataGridView3.Rows.Count - 1; i++)
            {
              //  MessageBox.Show(dataGridView3.Rows[i].Cells[2].Value.ToString());
                  if (checkBox5.Checked == false)
                 {
                     _PEList.Add((double.Parse(dataGridView3.Rows[i].Cells[6].Value.ToString().Trim()) -
                                     double.Parse(dataGridView3.Rows[i - 1].Cells[6].Value.ToString().Trim()))/
                                    ((double.Parse(dataGridView3.Rows[i].Cells[4].Value.ToString().Trim())) -
                                     (double.Parse(dataGridView3.Rows[i - 1].Cells[4].Value.ToString().Trim()))));
                 }
                if (checkBox5.Checked == true)
                {
                    if (i > 0)
                    {
                        //   MessageBox.Show(dataGridView3.Rows[i+1].Cells[6].Value + " " + (double)dataGridView3.Rows[i].Cells[6].Value + " " + dataGridView3.Rows[i+1].Cells[4].Value + " " + dataGridView3.Rows[i].Cells[4].Value);
                        _PEList.Add((double.Parse(dataGridView3.Rows[i].Cells[6].Value.ToString().Trim()) -
                                     double.Parse(dataGridView3.Rows[i - 1].Cells[6].Value.ToString().Trim()))/
                                    ((double.Parse(dataGridView3.Rows[i].Cells[4].Value.ToString().Trim())*3.75) -
                                     (double.Parse(dataGridView3.Rows[i - 1].Cells[4].Value.ToString().Trim())) * 3.75));
                    }
                }
            }

            for (int i = 0; i < _PEList.Count; i++)
            {
                dataGridView3.Rows[i + 1].Cells[7].Value = _PEList[i];
            }
        }



        private void chart1_AxisScrollBarClicked(object sender, ScrollBarEventArgs e)
        {

        }

        private void button13_Click(object sender, EventArgs e)
        {
            foreach (var item in chart1.Series)
            {
                MessageBox.Show(item.LegendText);
            }

        }

        private void comboBox2_SelectedIndexChanged_1(object sender, EventArgs e)
        {

            dataGridView4.Visible = false;

            this.dataGridView4.CellValueNeeded += new
                DataGridViewCellValueEventHandler(dataGridView41_CellValueNeeded);

            customers2.Clear();

            for (int k = 0; k < MyAllSensors.Count; k++)
            {
                if (comboBox2.Text == MyAllSensors[k].KKS_Name)
                {
                    for (int i = 0; i < MyAllSensors[k].MyListRecordsForOneKKS.Count; i++)
                    {
                        this.customers2.Add(
                            new MyVirtualClass(MyAllSensors[k].MyListRecordsForOneKKS[i].Value.ToString()));
                    }
                }
            }
            if (this.dataGridView4.RowCount == 0)
            {
                this.dataGridView4.RowCount = 1;
            }

            this.dataGridView4.RowCount = MyAllSensors[0].MyListRecordsForOneKKS.Count;
            dataGridView4.Visible = true;

            comboBox2.BackColor = Color.LightBlue;
        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            dataGridView4.Visible = false;

            this.dataGridView4.CellValueNeeded += new
                DataGridViewCellValueEventHandler(dataGridView42_CellValueNeeded);

            //   dataGridView4.Rows.Clear();
            customers3.Clear();


            for (int k = 0; k < MyAllSensors.Count; k++)
            {
                if (comboBox1.Text == MyAllSensors[k].KKS_Name)
                {
                    for (int i = 0; i < MyAllSensors[k].MyListRecordsForOneKKS.Count; i++)
                    {
                        this.customers3.Add(
                            new MyVirtualClass(MyAllSensors[k].MyListRecordsForOneKKS[i].Value.ToString()));
                    }
                }
            }
            if (this.dataGridView4.RowCount == 0)
            {
                this.dataGridView4.RowCount = 1;
            }

            this.dataGridView4.RowCount = MyAllSensors[0].MyListRecordsForOneKKS.Count;
            dataGridView4.Visible = true;

            comboBox1.BackColor = Color.LightBlue;
        }

        private void comboBox3_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            dataGridView4.Visible = false;

            this.dataGridView4.CellValueNeeded += new
                DataGridViewCellValueEventHandler(dataGridView43_CellValueNeeded);

            //   dataGridView4.Rows.Clear();
            customers4.Clear();


            for (int k = 0; k < MyAllSensors.Count; k++)
            {
                if (comboBox3.Text == MyAllSensors[k].KKS_Name)
                {
                    for (int i = 0; i < MyAllSensors[k].MyListRecordsForOneKKS.Count; i++)
                    {
                        this.customers4.Add(
                            new MyVirtualClass(MyAllSensors[k].MyListRecordsForOneKKS[i].Value.ToString()));
                    }
                }
            }
            if (this.dataGridView4.RowCount == 0)
            {
                this.dataGridView4.RowCount = 1;
            }

            this.dataGridView4.RowCount = MyAllSensors[0].MyListRecordsForOneKKS.Count;
            dataGridView4.Visible = true;

            comboBox3.BackColor = Color.LightBlue;
        }

        private void comboBox5_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            dataGridView4.Visible = false;

            this.dataGridView4.CellValueNeeded += new
                DataGridViewCellValueEventHandler(dataGridView44_CellValueNeeded);

            //   dataGridView4.Rows.Clear();
            customers5.Clear();


            for (int k = 0; k < MyAllSensors.Count; k++)
            {
                if (comboBox5.Text == MyAllSensors[k].KKS_Name)
                {
                    for (int i = 0; i < MyAllSensors[k].MyListRecordsForOneKKS.Count; i++)
                    {
                        this.customers5.Add(
                            new MyVirtualClass(MyAllSensors[k].MyListRecordsForOneKKS[i].Value.ToString()));
                    }
                }
            }
            if (this.dataGridView4.RowCount == 0)
            {
                this.dataGridView4.RowCount = 1;
            }

            this.dataGridView4.RowCount = MyAllSensors[0].MyListRecordsForOneKKS.Count;
            dataGridView4.Visible = true;

            comboBox5.BackColor = Color.LightBlue;
        }

        private void comboBox4_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            dataGridView4.Visible = false;

            this.dataGridView4.CellValueNeeded += new
                DataGridViewCellValueEventHandler(dataGridView45_CellValueNeeded);

            //   dataGridView4.Rows.Clear();
            customers6.Clear();


            for (int k = 0; k < MyAllSensors.Count; k++)
            {
                if (comboBox4.Text == MyAllSensors[k].KKS_Name)
                {
                    for (int i = 0; i < MyAllSensors[k].MyListRecordsForOneKKS.Count; i++)
                    {
                        this.customers6.Add(
                            new MyVirtualClass(MyAllSensors[k].MyListRecordsForOneKKS[i].Value.ToString()));
                    }
                }
            }
            if (this.dataGridView4.RowCount == 0)
            {
                this.dataGridView4.RowCount = 1;
            }

            this.dataGridView4.RowCount = MyAllSensors[0].MyListRecordsForOneKKS.Count;
            dataGridView4.Visible = true;

            comboBox4.BackColor = Color.LightBlue;
        }


        public struct pertubResult
        {

            public double aH;
            public double b;
            public double Ro;
            public double tau;

            public List<double> FF; //чисто для вывода 

            //величина невязки
            public double SS;
        }

        public static pertubResult Calc(double Tt, List<double> time, List<double> I, List<double> R, List<double> dH)
        {
            pertubResult myresult = new pertubResult();
            myresult.FF = new List<double>();
            //тут основной код для вычислений

            myresult.tau = Tt;

            int cnt = 0;

            double F, g, c, dT, sH, sHH, sF, sFF, sR, sFH, sRH, sRF;


            double SA, SB, SC, SD, SE;

            double Rm;
            double sS;


            F = 0;
            sH = 0;
            sHH = 0;
            sF = 0;
            sFF = 0;
            sFH = 0;
            sR = 0;
            sRH = 0;
            sRF = 0;

            myresult.FF.Add(0);

            for (int i = 1; i < time.Count; i++)
            {
                dT = time[i] - time[i - 1];
                if (dT == 0)
                {
                    myresult.FF.Add(0);
                    continue;
                }
                g = Math.Exp(-dT/Tt);
                c = (1 - g)*Tt/dT;
                F = F*g + (I[i] - 1)*(1 - c) + (I[i - 1] - 1)*(c - g);
                //TODO: ЧТО ТАКОЕ FF??????????
                myresult.FF.Add(F);



                sH += dH[i];
                sHH += dH[i]*dH[i];
                sF += F;
                sFF += F*F;
                sR += R[i];
                sFH += F*dH[i];
                sRH += R[i]*dH[i];
                sRF += R[i]*F;
                cnt++;
            }


            //считаем суммы

            SA = sH*sH/cnt/cnt - sHH/cnt;
            SB = sF*sH/cnt/cnt - sFH/cnt;
            SC = sR*sH/cnt/cnt - sRH/cnt;
            SD = sF*sF/cnt/cnt - sFF/cnt;
            SE = sR*sF/cnt/cnt - sRF/cnt;


            myresult.aH = (SC*SD - SB*SE)/(SA*SD - SB*SB);
           // MessageBox.Show(myresult.aH.ToString());
            myresult.b = (SA*SE - SB*SC)/(SA*SD - SB*SB);

            myresult.Ro = sR/cnt - myresult.aH*sH/cnt - myresult.b*sF/cnt;


            // невязки


            F = 0;
            cnt = 0;
            sS = 0;

            for (int i = 1; i < time.Count; i++)
            {
                dT = time[i] - time[i - 1];
                if (dT == 0)
                {
                    continue;
                }

                g = Math.Exp(-dT/Tt);
                c = (1 - g)*Tt/dT;
                F = F*g + (I[i] - 1)*(1 - c) + (I[i - 1] - 1)*(c - g);

                Rm = myresult.Ro + myresult.aH*dH[i] + myresult.b*F;
                sS += (R[i] - Rm)*(R[i] - Rm);

                cnt++;
            } //i
            myresult.SS = sS;
            return myresult;
        }
        //   tempR = new pertubResult();
           //     for (int i = 0; i < 400; i++)
            //    {
               //     tempR = Calc(3 + i / 200.0, _tList, _jKorList, _rCalcList, _dHList);
                //    if (tempR.SS > Ss)
                //    {
                //        MessageBox.Show(tempR.SS + " " + Ss);
                 //       tempR = Calc(3 + (i - 1) / 200.0, _tList, _jKorList, _rCalcList, _dHList);
                 //       break;
                //    }
                //    lll.WriteLine(tempR.aH + " " + tempR.tau + " " + tempR.SS);
                 //   Ss = tempR.SS;
              //  }
        private void button14_Click(object sender, EventArgs e)
        {
            if (saveFileDialog4.ShowDialog() == DialogResult.OK)
            {
                StreamWriter MyRecord = new StreamWriter(saveFileDialog4.FileName + ".csv", false, Encoding.Default);
                saveFileDialog4.Filter = "out data (*.csv)|*.csv|All files (*.*)|*.*";
                saveFileDialog4.DefaultExt = "csv";
            //    for (int i = 0; i < UPPER; i++)
            //    {
                    
               // }
            }
        }

        private void button13_Click_1(object sender, EventArgs e)
        {
            StreamReader r = new StreamReader("D:\\GitHub\\Differ\\Файл для заполнения таблицы для расчета ПЭ.txt",  Encoding.Default);
            string line = r.ReadLine();
     int _indexRow1 = 0;
         // dataGridView3.Rows.Add();
          while ((line = r.ReadLine()) != null)
           {
               dataGridView3.Rows.Add();
               for (int i = 0; i < 6; i++)
               {
                   dataGridView3.Rows[_indexRow].Cells[i].Value = line.Split('\t')[i].Trim();

              }
               _indexRow++;
          }

            r.Close();


          //  double rr = double.Parse(dataGridView3.Rows[0].Cells[1].Value.ToString().Trim());
        //   double gg = double.Parse(rr);            
         //  MessageBox.Show(gg.ToString());
    
        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {
            try
            {
                PE = double.Parse(textBox15.Text);
            }
            catch (Exception ex)
            {

                MessageBox.Show("Не число! Разделитель?");
            }
         
        }

        private void button16_Click(object sender, EventArgs e)
        {
            StreamWriter b = new StreamWriter("D:\\1.txt");
            for (int i = 0; i < _tList.Count; i++)
            {
                b.WriteLine(_tList[i] + "\t" + _rCalcList[i] + "\t" + _jExpList[i]*_jRealList[0]);
            }
            b.Close();
        }



    }
}
    

