﻿using System;
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
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System.Drawing.Printing;

//псевдонимы
using SD = System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using SmallPertubation;

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

        MyListOfSensors allSensors = new MyListOfSensors();


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
            }//
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
            allSensors.Clear();
            AllLEGENDS.Clear();
           
            //открытьToolStripMenuItem.Enabled = true;

            NumberSeriesNew = 0;
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

        int NumberSeriesNew = 0;

        List<string> AllLEGENDS = new List<string>();

        private void добавитьНаОсьXToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int j = 0; j < allSensors.Count; j++)
            {
                if ((string)checkedListBox1.Text.Split('\t')[0] == allSensors[j].KKS_Name)
                {
                    for (int i = 0; i < allSensors[j].MyListRecordsForOneKKS.Count; i++)
                    {
                       // MessageBox.Show(i + " " + allSensors[j].MyListRecordsForOneKKS[i].Value.ToString());
                     
                        chart1.Series[NumberSeriesNew].Points.AddXY(i,  allSensors[j].MyListRecordsForOneKKS[i].Value);
                    }
                    chart1.Series[NumberSeriesNew].IsVisibleInLegend = true;
                    chart1.Series[NumberSeriesNew].LegendText = allSensors[j].KKS_Name;
                    AllLEGENDS.Add(allSensors[j].KKS_Name);
                }
            }
          //  chart1.ChartAreas[0].AxisY.Maximum = 0.0215;
          //  chart1.ChartAreas[0].AxisY.Minimum = -0.0215;
            chart1.Series[NumberSeriesNew].BorderWidth = 3;
         //   chart1.Series[1].BorderWidth = 4;
    //        chart1.ChartAreas[0].AxisY.Maximum = 0.06;
      //      chart1.ChartAreas[0].AxisY.Minimum =- 0.06;
        //    Form1 main = this.Owner as Form1;
            button6.Enabled = true;
         //   MessageBox.Show(main.chart1.ChartAreas[0].AxisX.Maximum.ToString());
          //  MessageBox.Show(((double)chart1.ChartAreas[0].AxisX.Minimum).ToString());
            NumberSeriesNew++;
            chart1.ChartAreas[0].Position.Auto = true;
            //checkBox4.Enabled = true;
            //очиститьГрафикToolStripMenuItem.Enabled = true;
       //     checkBox5.Enabled = true;
          //  checkBox1.Enabled = true;
            // chart1.Series[0].XValueType = ChartValueType.Time;
        }

        double prosent;

        List<double> MaXes = new List<double>();
        List<double> MiNes = new List<double>();
        private void добавитьНаОсьYToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int j = 0; j < allSensors.Count; j++)
            {
                if ((string)checkedListBox1.Text.Split('\t')[0] == allSensors[j].KKS_Name)
                {
                    for (int i = 0; i < allSensors[j].MyListRecordsForOneKKS.Count; i++)
                    {
                        prosent = ((allSensors[j].MyListRecordsForOneKKS[i].Value - MiNes[j]) / (MaXes[j] - MiNes[j])) * 100;
                        chart1.Series[NumberSeriesNew].Points.AddXY(i, prosent);
                    }
                    chart1.Series[NumberSeriesNew].YAxisType = AxisType.Secondary;
                    chart1.Series[NumberSeriesNew].IsVisibleInLegend = true;
                    chart1.Series[NumberSeriesNew].LegendText = allSensors[j].KKS_Name;
                    AllLEGENDS.Add(allSensors[j].KKS_Name);
                }
            }
            NumberSeriesNew++;
            //checkBox4.Enabled = true;
            //очиститьГрафикToolStripMenuItem.Enabled = true;
       //     checkBox5.Enabled = true;
            //checkBox1.Enabled = true;

        //    chart1.ChartAreas[0].AxisY2.Minimum = 0;
          //  chart1.ChartAreas[0].AxisY2.Maximum = 100;

        }
        SD.DataTable TochkaDannih = new SD.DataTable();
        private void Form1_Load(object sender, EventArgs e)
        {
            
            chart1.ChartAreas[0].CursorX.IsUserEnabled = true;
            chart1.ChartAreas[0].AxisX.ScrollBar.IsPositionedInside = true;
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

            comboBox1.Text = "J1";
            comboBox2.Text = "R1";
            textBox3.BackColor = Color.Red;
            textBox4.BackColor = Color.Red;
            tabPage1.Text = "Зона обработки";
            tabPage2.Text = "Результаты";
            button6.Visible = true;
            button6.BackColor = Color.Yellow;
            button7.Visible = false;
            tabPage3.Text = "Параметры";
            tabPage4.Text = "Данные";
            button6.Enabled = false;
            chart1.ChartAreas[0].BackColor = Color.Gainsboro;
           
           for (int i = 0; i <chart1.Series.Count; i++)
               chart1.Series[i].IsVisibleInLegend = false;

            chart1.ChartAreas[0].AxisX.TitleFont = new System.Drawing.Font("Times New Roman", 14, FontStyle.Regular);
            chart1.ChartAreas[0].AxisY.TitleFont = new System.Drawing.Font("Times New Roman", 14, FontStyle.Regular);
            //button1.Enabled = false;
       

            //       правкаToolStripMenuItem.Visible = true;
     //    label24.Text = "d\u03F1" + "|dH";
        //    label25.Text = "d\u03F1" + "|dH";
            //  toolTip1.Show("Для начала откройте файл данных (Файл->Открыть)", this.menuStrip1);
          //  dataGridView4.DataSource = TochkaDannih;
            dataGridView4.Columns.Add("Время", "Время");
            dataGridView4.Columns.Add("I, A", "I, A");
            dataGridView4.Columns.Add("\u03F1, beff", "\u03F1, beff");
            dataGridView4.Columns.Add("H, %", "H, %");
            dataGridView4.Columns.Add("F", "F");



            dataGridView2.Columns.Add("\u03B1(H) \u03B2/см", "\u03B1(H) \u03B2/см");
          //  dataGridView2.Columns.Add("\u03B1(N)", "\u03B1(N)");
            dataGridView2.Columns.Add("По пичкам \u03B1(H) \u03B2/см", "По пичкам \u03B1(H) \u03B2/см");

            dataGridView4.Columns[0].Width = 200;
            dataGridView4.Columns[1].Width = 200;
            dataGridView4.Columns[2].Width = 200;
            dataGridView4.Columns[3].Width = 200;
            dataGridView4.Columns[4].Width = 200;



         



            this.dataGridView1.VirtualMode = true;
            //    this.dataGridView3.VirtualMode = true;

            DataGridViewTextBoxColumn companyNameColumn = new DataGridViewTextBoxColumn();
            companyNameColumn.HeaderText = "Время";
            companyNameColumn.Name = "Время";
            this.dataGridView1.Columns.Add(companyNameColumn);

            DataGridViewTextBoxColumn companyNameColumn1 = new DataGridViewTextBoxColumn();
            companyNameColumn1.HeaderText = "Значение";
            companyNameColumn1.Name = "Значение";
            this.dataGridView1.Columns.Add(companyNameColumn1);

            DataGridViewTextBoxColumn companyNameColumn2 = new DataGridViewTextBoxColumn();
            companyNameColumn2.HeaderText = "Время";
            companyNameColumn2.Name = "Время";
            //   this.dataGridView3.Columns.Add(companyNameColumn2);

            DataGridViewTextBoxColumn companyNameColumn3 = new DataGridViewTextBoxColumn();
            companyNameColumn3.HeaderText = "Значение";
            companyNameColumn3.Name = "Значение";
            //  this.dataGridView3.Columns.Add(companyNameColumn3);





            dataGridView1.Columns[0].Width = 300;
            dataGridView1.Columns[1].Width = 300;
            
          
        }
        private void dataGridView3_CellValueNeeded(object sender,
     System.Windows.Forms.DataGridViewCellValueEventArgs e)
        {
            //this.dataGridView1.RowCount = 1;
            // If this is the row for new records, no values are needed.
            //  if (e.RowIndex == this.dataGridView3.RowCount - 1) return;
            //   if (e.RowIndex == this.dataGridView3.RowCount - 1) return;

            MyVirtualClass customerTmp = null;

            // Store a reference to the Customer object for the row being painted.
            if (e.RowIndex == rowInEdit)
            {
                customerTmp = this.customerInEdit;
            }
            else
            {
                customerTmp = (MyVirtualClass)this.customers[e.RowIndex];
            }

            // Set the cell value to paint using the Customer object retrieved.
            //switch (this.dataGridView3.Columns[e.ColumnIndex].Name)
            //{
            //    case "Время":
            //        e.Value = customerTmp.time;
            //        break;

            //    case "Значение":
            //        e.Value = customerTmp.value;
            //        break;
            //}
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
                customerTmp = (MyVirtualClass)this.customers[e.RowIndex];
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
        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.Visible = false;
                //   Sencors myOneKKS = new Sencors();
                Sencors myOneKKS = new Sencors();
                //e.Node.Parent.Index
                //  List<double> MyParametr = new List<double>();
                //    MessageBox.Show(MyListSensors.Count.ToString());
                //    MessageBox.Show(MyListSensors[0].Count.ToString() + " " + MyListSensors[1].Count.ToString());
             
                    myOneKKS = allSensors.getOneKKSByIndex(checkedListBox1.SelectedIndex);

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
                    this.customers.Add(new MyVirtualClass(myOneKKS.MyListRecordsForOneKKS[i].DateTime.ToString("HH:mm:ss.fff"), myOneKKS.MyListRecordsForOneKKS[i].Value.ToString()));

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
        TextAnnotation text1 =  new TextAnnotation();
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

            for (int ii = 0; ii < NumberSeriesNew; ii++)
            {
                sum = 0;
                aver = 0;
                for (int j = 0; j < allSensors.Count; j++)
                {
                    if (chart1.Series[ii].LegendText == allSensors[j].KKS_Name && (int)chart1.ChartAreas[0].CursorX.Position > 0)
                    {
                        for (int i = 0; i < NumberPoints; i++)
                        {
                            sum = sum + allSensors[j].MyListRecordsForOneKKS[(int)chart1.ChartAreas[0].CursorX.Position - i].Value;
                        }
                        aver = sum / NumberPoints;
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
                if (dt.GetDataPresent(typeof(string)))
                {
                    string tb = (string)(dt.GetData(typeof(string)));
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
                if (((DataGridView)sender).SelectedCells.Count > 0)
                {
                    copySelectedRowsToClipboard((DataGridView)sender);
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
        List<double> MAXnew = new List<double>();
        List<double> MINnew = new List<double>();
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
                if ((string)checkedListBox1.Text.Split('\t')[0] == AllLEGENDS[j])
                {
                    indexGrafic = j;
                    chart1.Series[indexGrafic].Points.Clear();
                    {
                        for (int k = 0; k < allSensors.Count; k++)
                        {
                            if (AllLEGENDS[j] == allSensors[k].KKS_Name)
                            {
                                for (int i = 0; i < allSensors[k].MyListRecordsForOneKKS.Count; i++)
                                {
                                    prosent = ((allSensors[k].MyListRecordsForOneKKS[i].Value - MINnew[k]) / (MAXnew[k] - MINnew[k])) * 100;
                                    chart1.Series[indexGrafic].Points.AddXY(i, prosent);
                                }
                                chart1.Series[indexGrafic].IsVisibleInLegend = true;
                                chart1.Series[indexGrafic].LegendText = allSensors[k].KKS_Name;
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
                if (((DataGridView)sender).SelectedCells.Count > 0)
                {
                    copySelectedRowsToClipboard((DataGridView)sender);
                }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            chart1.ChartAreas[0].CursorX.Position--;
            //MessageBox.Show(chart1.ChartAreas[0].CursorX.Position.ToString());


            for (int ii = 0; ii < NumberSeriesNew; ii++)
            {//колличество нарисованных сириосов тоесть колличество выводимых параметров
             //   TochkaDannih.Columns.Add(chart1.Series[ii].LegendText);
                for (int j = 0; j < allSensors.Count; j++)
                {
                    if (chart1.Series[ii].LegendText == allSensors[j].KKS_Name)
                    {
                      //  DateTime WindowsTime = new DateTime(1970, 1, 1).AddSeconds(allSensors[j].MyListRecordsForOneKKS[(int)chart1.ChartAreas[0].CursorX.Position].DateTime);
                     //   dataGridView3.Rows[0].Cells[0].Value = WindowsTime.ToString("HH:mm:ss");

                  
                        textBox3.Text = allSensors[j].MyListRecordsForOneKKS[(int)chart1.ChartAreas[0].CursorX.Position].ValueTimeForDAT.ToString();
                        textBox4.Text = allSensors[j].MyListRecordsForOneKKS[(int)chart1.ChartAreas[0].CursorX.Position].DateTime.ToString("HH:mm:ss");
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
            button6.Enabled = true;
         //   Time_Per.Clear();
        //    Time_Per.RemoveAt(Time_Per.Count-1);

            chart1.ChartAreas[0].CursorX.Position--;

            //  indexes.Add(position);
            //кол/личество нарисованных сириосов тоесть колличество выводимых параметров
            try
            {
                for (int j = 0; j < allSensors.Count; j++)
                {
                    if (comboBox2.Text == allSensors[j].KKS_Name)
                    {
                        chart1.Series[19].ChartType = SeriesChartType.Point;
                        chart1.Series[19].Color = Color.Black;
                        chart1.Series[19].Points.Clear();
                        position = (int)chart1.ChartAreas[0].CursorX.Position;

                        //MessageBox.Show(((int)chart1.ChartAreas[0].CursorX.Position).ToString());
                        textBox3.Text = allSensors[j].MyListRecordsForOneKKS[position].ValueTimeForDAT.ToString();
                        textBox4.Text = allSensors[j].MyListRecordsForOneKKS[position].DateTime.ToString("HH:mm:ss");
                        DataPoint dp = new DataPoint(chart1.ChartAreas[0].CursorX.Position, allSensors[j].MyListRecordsForOneKKS[position].Value);
                        dp.MarkerStyle = MarkerStyle.Cross;
                        dp.MarkerSize = 10;
                        dp.IsValueShownAsLabel = true;
                        chart1.Series[19].Points.Add(dp);
                    }
                }
                //  chart1.ChartAreas[0].CursorX.Position++;
                position = (int)chart1.ChartAreas[0].CursorX.Position;

            }
            catch
            {
                chart1.ChartAreas[0].CursorX.Position = position;
            }
        }
        List<int> indexes = new List<int>();
        List<int> BeginIndex = new List<int>();
        List<int> EndIndex = new List<int>();
        int position;
        List<double> Time_Per = new List<double>();
        List<double> myTok = new List<double>();
        double beg_per;
        double end_per;
     //   myOneVozmuchenie.N_Per=0;
        private void button5_Click(object sender, EventArgs e)
        {
            button6.Enabled = true;
       //     Time_Per.Add(allSensors[0].MyListRecordsForOneKKS[(int)chart1.ChartAreas[0].CursorX.Position].value1);
             chart1.ChartAreas[0].CursorX.Position++;

        //кол/личество нарисованных сириосов тоесть колличество выводимых параметров
            try
            {
                for (int j = 0; j < allSensors.Count; j++)
                {
                    if (comboBox2.Text == allSensors[j].KKS_Name)
                    {
                        chart1.Series[19].ChartType = SeriesChartType.Point;
                        chart1.Series[19].Color = Color.Black;
                        chart1.Series[19].Points.Clear();
                        position = (int)chart1.ChartAreas[0].CursorX.Position;

                        //MessageBox.Show(((int)chart1.ChartAreas[0].CursorX.Position).ToString());
                        textBox3.Text = allSensors[j].MyListRecordsForOneKKS[position].ValueTimeForDAT.ToString();
                        textBox4.Text = allSensors[j].MyListRecordsForOneKKS[position].DateTime.ToString("HH:mm:ss");
                        DataPoint dp = new DataPoint(chart1.ChartAreas[0].CursorX.Position, allSensors[j].MyListRecordsForOneKKS[position].Value);
                        dp.MarkerStyle = MarkerStyle.Cross;
                        dp.MarkerSize = 10;
                        dp.IsValueShownAsLabel = true;
                        chart1.Series[19].Points.Add(dp);
                    }
                }
              //  chart1.ChartAreas[0].CursorX.Position++;
                position = (int)chart1.ChartAreas[0].CursorX.Position;
     
            }
            catch
            {
                chart1.ChartAreas[0].CursorX.Position = position;
            }
            
            //if (VozmSeries.Points.Count == 0)
            //{
            //    ///найдем индекс
            //    ///
            //    this.index = 0;
            //    do
            //    {
            //        index++;
            //    } while (this.form1Chart1.ChartAreas[0].CursorX.Position > this.myInputData.Data[0][index]);
            //    this.VozmSeries.Points.AddXY(this.myInputData.Data[0][index], this.myInputData.Data[MyConst.r1][index]);
            //}
            //else
            //{
            //    index++;
            //    //удалим последнюю точку
            //    this.VozmSeries.Points.RemoveAt(this.VozmSeries.Points.Count - 1);
            //    this.VozmSeries.Points.AddXY(this.myInputData.Data[0][index], this.myInputData.Data[MyConst.r1][index]);
            //    //this.FixVozmSeries.Points.AddXY(this.myData.Data[0][i], this.myData.Data[MyConst.r1][i]);
            //}
        }

        private void sddsToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
        int index;
        int IndexReac()
        {
            index = 0;
            for (int i = 0; i < allSensors.Count; i++)
            {
                if (comboBox2.Text == allSensors[i].KKS_Name)
                {
                    index = i;
                }
            }
            return index;
        }
        private void chart1_MouseDown_1(object sender, MouseEventArgs e)
        {
        //    IndexReac();
            if (e.Button == MouseButtons.Right)
            {
                button6.Enabled = true;
                chart1.ContextMenuStrip = contextMenuStrip2;
            }
            if (e.Button == MouseButtons.Left)
            {
                   try
            {
                for (int j = 0; j < allSensors.Count; j++)
                {
                    if (comboBox2.Text == allSensors[j].KKS_Name)
                    {
                        chart1.Series[19].ChartType = SeriesChartType.Point;
                        chart1.Series[19].Color = Color.Black;
                        chart1.Series[19].Points.Clear();
                        position = (int)chart1.ChartAreas[0].CursorX.Position;

                        //MessageBox.Show(((int)chart1.ChartAreas[0].CursorX.Position).ToString());
                        textBox3.Text = allSensors[j].MyListRecordsForOneKKS[position].ValueTimeForDAT.ToString();
                        textBox4.Text = allSensors[j].MyListRecordsForOneKKS[position].DateTime.ToString("HH:mm:ss");
                        DataPoint dp = new DataPoint(chart1.ChartAreas[0].CursorX.Position, allSensors[j].MyListRecordsForOneKKS[position].Value);
                        dp.MarkerStyle = MarkerStyle.Cross;
                        dp.MarkerSize = 10;
                        dp.IsValueShownAsLabel = true;
                        chart1.Series[19].Points.Add(dp);
                    }
                }
              //  chart1.ChartAreas[0].CursorX.Position++;
                position = (int)chart1.ChartAreas[0].CursorX.Position;
     
            }
            catch
            {
                chart1.ChartAreas[0].CursorX.Position = position;
            }
            }

        }
        bool ter = false;
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
                chart1.ChartAreas[0].AxisY.LabelAutoFitMinFontSize = (int)numericUpDown4.Value;
            }
            catch
            {

            }
        }

        private void numericUpDown6_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                chart1.ChartAreas[0].AxisY2.LabelAutoFitMinFontSize = (int)numericUpDown6.Value;
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
                chart1.Titles[0].Font = new System.Drawing.Font("Times New Roman", (int)numericUpDown1.Value, FontStyle.Regular);
            }
            catch
            {

            }
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {

            try
            {
                chart1.ChartAreas[0].AxisX.TitleFont = new System.Drawing.Font("Times New Roman", (int)numericUpDown2.Value, FontStyle.Regular);
            }
            catch
            {

            }
        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {

            try
            {
                chart1.ChartAreas[0].AxisY.TitleFont = new System.Drawing.Font("Times New Roman", (int)numericUpDown3.Value, FontStyle.Regular);
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
        List<int> ForERORR = new List<int>();
        private void button6_Click(object sender, EventArgs e)
        {
            ForERORR.Add(0);
            if (BeginIndex.Count == 0)
            {
                pdH.Add(0);
            }
            //   Time_Per.Add(allSensors[0].MyListRecordsForOneKKS[(int)chart1.ChartAreas[0].CursorX.Position].value1);

            //   MessageBox.Show("gv");



            BeginIndex.Add(position);

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
                button7.BackColor = Color.Yellow;
                // No_Perem_H();
                //  MessageBox.Show(BeginIndex[0].ToString());
                // myOneVozmuchenie.N_Per = 0;
                //       MessageBox.Show(MyTok.Count.ToString() + " " + myReactivity.Count.ToString() + " " + pdH.Count.ToString());
            }
        }
        private void button7_Click(object sender, EventArgs e)
        {
            EndIndex.Add(position);
            ForERORR.Add(1);
         //  if(EndIndex.Count == 1)
           // {
            button8.Enabled = true;             //   Time_Per.RemoveAt(1);
           // }
           // Razdroblenie_H();
           // Time_Per.Clear();
            
            button6.Visible = true;
            button7.Visible = false;
            button6.BackColor = Color.Yellow;
       //     MessageBox.Show(EndIndex[EndIndex.Count-1].ToString());
         //  MessageBox.Show(myOneVozmuchenie.N_Per.ToString());
         //  MessageBox.Show(Time_Per.Count.ToString());
        }
      //  List<double> MyTok = new List<double>();
       // List<double> MyTime = new List<double>();
        List<double> proba = new List<double>();
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }
        int NOper = -2;
        private OneVozmuchenieData myOneVozmuchenie = new OneVozmuchenieData();
        List<double> pdH = new List<double>();
        List<double> myReactivity = new List<double>();
        double step;
     
   
        List<double> time = new List<double>();
        private void ReturnTokAndReactivity()
        { 
            for (int i = 0; i < allSensors.Count; i++)
            {
                if (comboBox1.Text == allSensors[i].KKS_Name)
                {
                    for (int j = BeginIndex[0]; j < EndIndex[EndIndex.Count-1]+1; j++)
                    {
                        myTok.Add(allSensors[i].MyListRecordsForOneKKS[j].Value / allSensors[i].MyListRecordsForOneKKS[BeginIndex[0]].Value);
                        time.Add(allSensors[i].MyListRecordsForOneKKS[j].ValueTimeForDAT);
                    }
                }
            }

            for (int i = 0; i < allSensors.Count; i++)
            {
                if (comboBox2.Text == allSensors[i].KKS_Name)
                {
                    for (int j = BeginIndex[0]; j < EndIndex[EndIndex.Count-1]+1; j++)
                    {
                        myReactivity.Add(allSensors[i].MyListRecordsForOneKKS[j].Value);
                        chart1.Series[10].Points.AddXY(j, allSensors[i].MyListRecordsForOneKKS[j].Value);
                    }

                }
            }
            double ddH = 0;
            step = 0;
            for (int j = 0; j < EndIndex.Count; j++)
            {
                for (int i = BeginIndex[j]; i < EndIndex[j]; i++)
                {
                    ddH -= (2 / (allSensors[0].MyListRecordsForOneKKS[EndIndex[j]].ValueTimeForDAT - allSensors[0].MyListRecordsForOneKKS[BeginIndex[j]].ValueTimeForDAT)) * (allSensors[0].MyListRecordsForOneKKS[i + 1].ValueTimeForDAT - allSensors[0].MyListRecordsForOneKKS[i].ValueTimeForDAT);
                    pdH.Add(ddH);                 
                }
                if (j != EndIndex.Count - 1)
                {
                    int k = j + 1;
                    for (int i = EndIndex[j]; i < BeginIndex[k]; i++)
                    {
                        pdH.Add(ddH);
                    }
                }     
            }
            chart1.Series[10].ChartType = SeriesChartType.Point;
            chart1.Series[NumberSeriesNew].BorderWidth = 3;
           
   //         label24.Text ="Просто: dp/dH = " + dRdH() + " beff/см";
           
            pertubResult t = new pertubResult();
            t = Search_a_b_p(time, myReactivity, pdH, myTok);
            //Search_a_b_p(time, myReactivity, pdH,myTok, time.Count, );



           /// MessageBox.Show(myTok.Count.ToString() + " " + myReactivity.Count.ToString() + " " + pdH.Count.ToString() + " " + t.FF.Count());
         //   label25.Text = "Сложно: dp/dH = " + t.aH.ToString() + " beff/см";
         //   if (dataGridView4.RowCount > 1)
           //     dataGridView4.Rows.Clear();
            Random random = new Random();
            int color = random.Next(255);
            for (int i = 0; i <pdH.Count; i++)
            {
                dataGridView4.Rows.Add(time[i], myTok[i], myReactivity[i], pdH[i], t.FF[i]);
                dataGridView4.Rows[i].DefaultCellStyle.BackColor = Color.Turquoise;
            }

            //sr.Write(myOneVozmuchenie.Nt);

            //sr.Write(";" + tempR.b);

            //sr.Write(";" + tempR.aH);
            //sr.Write(";" + tempR.b / myOneVozmuchenie.Nt);
            //sr.Write(";" + tempR.tau);
            //sr.Write(";" + tempR.Ro);
          
      //      MessageBox.Show(t.aH.ToString());

            dataGridView2.Rows.Add(t.aH, dRdH());
            dataGridView2.Rows[0].DefaultCellStyle.BackColor = Color.Yellow;
            BeginIndex.Clear();
            EndIndex.Clear();
            myTok.Clear();
            time.Clear();
            t.FF.Clear();
            myReactivity.Clear();
            pdH.Clear();
        }
        private double dRdH()
        {
            double DroAver = 0;
            for (int i = 0; i < allSensors.Count; i++)
            {
                if (comboBox2.Text == allSensors[i].KKS_Name)
                {
                    for (int j = 0; j < BeginIndex.Count; j++)
                    {

                        DroAver = DroAver + (allSensors[i].MyListRecordsForOneKKS[EndIndex[j]].Value - allSensors[i].MyListRecordsForOneKKS[BeginIndex[j]].Value);


                        //ведь каждый шаг должен в сумме давать 2
                        //for (int j = myOneVozmuchenie.PerIndex[i, 0] + 1; j <= myOneVozmuchenie.PerIndex[i, 1]; j++)
                        //{
                        //    ddH -= (2 / (myOneVozmuchenie.Per[i, 1] - myOneVozmuchenie.Per[i, 0])) * (myInputData.Data[0][j] - myInputData.Data[0][j - 1]);
                        //    pdH.Add(ddH);
                        //}


                    }
                    DroAver = DroAver / BeginIndex.Count;
                    DroAver = DroAver / -2;
                    DroAver = DroAver * MyConst.Rect.Beff;
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
        private pertubResult Search_a_b_p(List<double> MyTimeInterval, List<double> MyReactivity, List<double> dH, List<double> TOK)
        {
            double Ss = 1000;
            pertubResult tempR = new pertubResult();

            for (int i = 0; i < 400; i++)
            {
                tempR = tempR.Calc(3 + i / 200, time, TOK, MyReactivity, dH);
                if (tempR.SS > Ss)
                {
                    tempR = tempR.Calc(3 + (i - 1) / 200, time, TOK, MyReactivity, pdH);
                    break;
                }
                Ss = tempR.SS;
            }
            tempR.Ro = tempR.Ro * MyConst.Rect.Beff;
            tempR.aH = tempR.aH * MyConst.Rect.Beff;
            tempR.b = tempR.b * MyConst.Rect.Beff;
            tempR.SS = tempR.SS * MyConst.Rect.Beff;
            return tempR;
        }
        private void button8_Click(object sender, EventArgs e)
        {
            dataGridView4.Rows.Clear();
            dataGridView2.Rows.Clear();
            if (BeginIndex.Count == EndIndex.Count)
            {
                ReturnTokAndReactivity();
            }
            if(BeginIndex.Count != EndIndex.Count)
            {
                MessageBox.Show("Введите конечную точку!");
            }

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
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                foreach (var item in openFileDialog1.FileNames)
                {

                    allSensors.LoadFromFile(item, allSensors);

                }
            }


            foreach (Sencors item in allSensors)
            {
                checkedListBox1.Items.Add(item.KKS_Name);
                comboBox1.Items.Add(item.KKS_Name);
                comboBox2.Items.Add(item.KKS_Name);
                comboBox3.Items.Add(item.KKS_Name);
                comboBox4.Items.Add(item.KKS_Name);
                comboBox5.Items.Add(item.KKS_Name);
            }
            
        }
        bool flag = true;
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
            try
            {
                NumberSeriesNew = 0;

                chart1.Titles.Clear();
           //     добавитьНаОсьXToolStripMenuItem.Enabled = false;

             //   checkBox4.Enabled = true;
                //  открытьToolStripMenuItem.Enabled = true;
                // checkedListBox1.Items.Clear();
                checkedListBox1.Items.Clear();
                //   MyAllSensors.Clear();
                allSensors.Clear();
                chart1.Legends.Clear();
                //   checkedListBox1.Items.Clear();
             //   chart1.Series.Clear();
                chart1.ChartAreas[0].Position.Auto = true;

                for (int i = 0; i < chart1.Series.Count; i++)
                {
                    chart1.Series[i].IsVisibleInLegend = false;
                    chart1.Series[i].Points.Clear();
                }
            }
            catch (Exception ex)
            {

            }
        }
    }
    public struct pertubResult
    {

        public double aH;
        public double b;
        public double Ro;
        public double tau;

        public List<double> FF;  //чисто для вывода 

        //величина невязки
        public double SS;






        //List<double> time = new List<double>();
        //List<double> I = new List<double>();
        //List<double> P = new List<double>();
        //List<double> pdH = new List<double>();

        /// <summary>
        /// основная функция для расчета
        /// </summary>
        /// <param name="Tau">постоянная времени разогрева</param>
        /// <param name="time">массив времени</param>
        /// <param name="I">массив токов</param>
        /// <param name="P">массив реактивностей</param>
        /// <param name="dH">массив перемещений группы</param>
        /// <returns></returns>
        /// 

        public pertubResult Calc(double Tt, List<double> time, List<double> I, List<double> R, List<double> dH)
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
                g = Math.Exp(-dT / Tt);
                c = (1 - g) * Tt / dT;
                F = F * g + (I[i] - 1) * (1 - c) + (I[i - 1] - 1) * (c - g);
                myresult.FF.Add(F);



                sH += dH[i];
                sHH += dH[i] * dH[i];
                sF += F;
                sFF += F * F;
                sR += R[i];
                sFH += F * dH[i];
                sRH += R[i] * dH[i];
                sRF += R[i] * F;
                cnt++;
            }


            //считаем суммы

            SA = sH * sH / cnt / cnt - sHH / cnt;
            SB = sF * sH / cnt / cnt - sFH / cnt;
            SC = sR * sH / cnt / cnt - sRH / cnt;
            SD = sF * sF / cnt / cnt - sFF / cnt;
            SE = sR * sF / cnt / cnt - sRF / cnt;


            myresult.aH = (SC * SD - SB * SE) / (SA * SD - SB * SB);
            myresult.b = (SA * SE - SB * SC) / (SA * SD - SB * SB);

            myresult.Ro = sR / cnt - myresult.aH * sH / cnt - myresult.b * sF / cnt;


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

                g = Math.Exp(-dT / Tt);
                c = (1 - g) * Tt / dT;
                F = F * g + (I[i] - 1) * (1 - c) + (I[i - 1] - 1) * (c - g);

                Rm = myresult.Ro + myresult.aH * dH[i] + myresult.b * F;
                sS += (R[i] - Rm) * (R[i] - Rm);

                cnt++;
            } //i
            myresult.SS = sS;



            return myresult;
        }


    }
}
