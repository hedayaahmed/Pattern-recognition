using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace Project
{
    public partial class Form1 : Form
    {
        public class DataSample
        {
            public List<double> Xs = new List<double>();
            public double Class;
        }

        public class Distance
        {
            public double distance;
            public double Class;
        }

        List<DataSample> DataSamples = new List<DataSample>();
        List<Distance> Distances = new List<Distance>();

        List<Distance> SortedD = new List<Distance>();
        int K, M;
        int Cl;

        DataSample ptrav = new DataSample();

        public Form1()
        {
            InitializeComponent();
            this.Load += new EventHandler(Form1_Load);
        }

        void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.ReadOnly = true;
            dataGridView1.RowCount = 215;
            dataGridView1.ColumnCount = 11;

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"D:\CS487 Pattern\Project\Project\bin\Debug\4.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            for (int i = 2; i < 216; i++)
            {
                DataSample pnn = new DataSample();
                int j;
                for (j = 2; j <= colCount; j++)
                {
                    pnn.Xs.Add(Convert.ToDouble(xlWorksheet.Cells[i, j].Value2));
                }
                pnn.Class = Convert.ToDouble(xlWorksheet.Cells[i, j - 1].Value2);
                DataSamples.Add(pnn);
            }

            for (int i = 0; i < 214; i++)
            {
                for (int j = 0; j < DataSamples[0].Xs.Count; j++)
                {
                    dataGridView1[j, i].Value = DataSamples[i].Xs[j].ToString();
                }
            }
            //MessageBox.Show(DataSamples[0].Xs.Count+"");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            K = Int32.Parse(textBox10.Text);
            ptrav.Xs.Clear();
            Distances.Clear();
            SortedD.Clear();
            Cl = '\0';
            ptrav.Xs.Add(double.Parse(textBox1.Text));
            ptrav.Xs.Add(double.Parse(textBox2.Text));
            ptrav.Xs.Add(double.Parse(textBox3.Text));
            ptrav.Xs.Add(double.Parse(textBox4.Text));
            ptrav.Xs.Add(double.Parse(textBox5.Text));
            ptrav.Xs.Add(double.Parse(textBox6.Text));
            ptrav.Xs.Add(double.Parse(textBox7.Text));
            ptrav.Xs.Add(double.Parse(textBox8.Text));
            ptrav.Xs.Add(double.Parse(textBox9.Text));

            for (int i = 0; i < 214; i++)
            {
                Distance pnn = new Distance();
                pnn.distance = 0;
                for (int j = 0; j < DataSamples[0].Xs.Count - 1; j++)
                {
                    pnn.distance += Math.Pow(ptrav.Xs[j] - DataSamples[i].Xs[j], 2);
                }
                pnn.Class = DataSamples[i].Class;
                Distances.Add(pnn);
            }

            SortedD = Distances.OrderBy(o => o.distance).ToList();
            double[] Count = new double[6];
            for (int i = 0; i < 6; i++)
            {
                Count[i] = 0;
            }

            for (int i = 0; i < K; i++)
            {
                if (SortedD[i].Class == 1)
                {
                    Count[0]++;
                }
                if (SortedD[i].Class == 2)
                {
                    Count[1]++;
                }
                if (SortedD[i].Class == 3)
                {
                    Count[2]++;
                }

                if (SortedD[i].Class == 5)
                {
                    Count[3]++;
                }
                if (SortedD[i].Class == 6)
                {
                    Count[4]++;
                }
                if (SortedD[i].Class == 7)
                {
                    Count[5]++;
                }
            }

            double max = -999999999;

            for (int i = 0; i < 6; i++)
            {
                if (Count[i] > max)
                {
                    max = Count[i];
                    if (i < 3)
                    {
                        Cl = i + 1;
                    }
                    else
                    {
                        Cl = i + 2;
                    }
                }
            }
            textBox12.Text = Cl.ToString();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            K = Int32.Parse(textBox10.Text);
            M = Int32.Parse(textBox11.Text);
            ptrav.Xs.Clear();
            Distances.Clear();
            SortedD.Clear();
            Cl = '\0';
            ptrav.Xs.Add(double.Parse(textBox1.Text));
            ptrav.Xs.Add(double.Parse(textBox2.Text));
            ptrav.Xs.Add(double.Parse(textBox3.Text));
            ptrav.Xs.Add(double.Parse(textBox4.Text));
            ptrav.Xs.Add(double.Parse(textBox5.Text));
            ptrav.Xs.Add(double.Parse(textBox6.Text));
            ptrav.Xs.Add(double.Parse(textBox7.Text));
            ptrav.Xs.Add(double.Parse(textBox8.Text));
            ptrav.Xs.Add(double.Parse(textBox9.Text));

            for (int i = 0; i < 214; i++)
            {
                Distance pnn = new Distance();
                pnn.distance = 0;
                for (int j = 0; j < DataSamples[0].Xs.Count-1; j++)
                {
                    pnn.distance += Math.Pow(ptrav.Xs[j] - DataSamples[i].Xs[j], 2);
                }
                pnn.Class = DataSamples[i].Class;
                Distances.Add(pnn);
            }

            //MessageBox.Show(Distances[0].Class + " ", Distances[0].distance + "");
            SortedD = Distances.OrderBy(o => o.distance).ToList();
            double[] MCalc = new double[6];
            double tmpv1 = 0;
            double tmpv2 = 0;
            for (int i = 0; i < K; i++)
            {
                if (SortedD[i].Class == 1)
                {
                    tmpv1 += 1 / Math.Pow(SortedD[i].distance, M);
                }
                tmpv2 += 1 / Math.Pow(SortedD[i].distance, M);
            }
            MCalc[0] = tmpv1 / tmpv2;

            tmpv1 = 0;
            tmpv2 = 0;
            for (int i = 0; i < K; i++)
            {
                if (SortedD[i].Class == 2)
                {
                    tmpv1 += 1 / Math.Pow(SortedD[i].distance, M);
                }
                tmpv2 += 1 / Math.Pow(SortedD[i].distance, M);
            }
            MCalc[1] = tmpv1 / tmpv2;

            tmpv1 = 0;
            tmpv2 = 0;
            for (int i = 0; i < K; i++)
            {
                if (SortedD[i].Class == 3)
                {
                    tmpv1 += 1 / Math.Pow(SortedD[i].distance, M);
                }
                tmpv2 += 1 / Math.Pow(SortedD[i].distance, M);
            }
            MCalc[2] = tmpv1 / tmpv2;

            tmpv1 = 0;
            tmpv2 = 0;
            for (int i = 0; i < K; i++)
            {
                if (SortedD[i].Class == 5)
                {
                    tmpv1 += 1 / Math.Pow(SortedD[i].distance, M);
                }
                tmpv2 += 1 / Math.Pow(SortedD[i].distance, M);
            }
            MCalc[3] = tmpv1 / tmpv2;

            tmpv1 = 0;
            tmpv2 = 0;
            for (int i = 0; i < K; i++)
            {
                if (SortedD[i].Class == 6)
                {
                    tmpv1 += 1 / Math.Pow(SortedD[i].distance, M);
                }
                tmpv2 += 1 / Math.Pow(SortedD[i].distance, M);
            }
            MCalc[4] = tmpv1 / tmpv2;

            tmpv1 = 0;
            tmpv2 = 0;
            for (int i = 0; i < K; i++)
            {
                if (SortedD[i].Class == 7)
                {
                    tmpv1 += 1 / Math.Pow(SortedD[i].distance, M);
                }
                tmpv2 += 1 / Math.Pow(SortedD[i].distance, M);
            }
            MCalc[5] = tmpv1 / tmpv2;
            double max=-999999999;
            
            for (int i = 0; i < 6; i++)
            {
                if (MCalc[i] > max)
                {
                    max = MCalc[i];
                    if (i < 3)
                    {
                        Cl = i + 1;
                    }
                    else
                    {
                        Cl = i + 2;
                    }
                }
            }
            textBox14.Text = Cl.ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            K = Int32.Parse(textBox10.Text);
            ptrav.Xs.Clear();
            Distances.Clear();
            SortedD.Clear();
            Cl = '\0';
            ptrav.Xs.Add(double.Parse(textBox1.Text));
            ptrav.Xs.Add(double.Parse(textBox2.Text));
            ptrav.Xs.Add(double.Parse(textBox3.Text));
            ptrav.Xs.Add(double.Parse(textBox4.Text));
            ptrav.Xs.Add(double.Parse(textBox5.Text));
            ptrav.Xs.Add(double.Parse(textBox6.Text));
            ptrav.Xs.Add(double.Parse(textBox7.Text));
            ptrav.Xs.Add(double.Parse(textBox8.Text));
            ptrav.Xs.Add(double.Parse(textBox9.Text));

            for (int i = 0; i < 214; i++)
            {
                Distance pnn = new Distance();
                pnn.distance = 0;
                for (int j = 0; j < DataSamples[0].Xs.Count - 1; j++)
                {
                    pnn.distance += Math.Pow(ptrav.Xs[j] - DataSamples[i].Xs[j], 2);
                }
                pnn.Class = DataSamples[i].Class;
                Distances.Add(pnn);
            }

            SortedD = Distances.OrderBy(o => o.distance).ToList();

            double max = -999999;
            double min = 999999;
            for (int i= 0; i < K; i++)
            {
                if(SortedD[i].distance>max)
                {
                    max = SortedD[i].distance;
                }
                if(SortedD[i].distance<min)
                {
                    min = SortedD[i].distance;
                }

            }

            double[] Weights = new double[6];
            for (int i = 0; i < 6; i++)
            {
                Weights[i] = 0;
            }

            for (int i = 0; i < K; i++)
            {
                if (SortedD[i].Class == 1)
                {
                    Weights[0] += (max - SortedD[i].distance) / (max - min);
                }
                if (SortedD[i].Class == 2)
                {
                    Weights[1] += (max - SortedD[i].distance) / (max - min);
                }
                if (SortedD[i].Class == 3)
                {
                    Weights[2] += (max - SortedD[i].distance) / (max - min);
                }
                if (SortedD[i].Class == 5)
                {
                    Weights[3] += (max - SortedD[i].distance) / (max - min);
                }
                if (SortedD[i].Class == 6)
                {
                    Weights[4] += (max - SortedD[i].distance) / (max - min);
                }
                if (SortedD[i].Class == 7)
                {
                    Weights[5] += (max - SortedD[i].distance) / (max - min);
                }
            }

            max = -999999;
            for (int i = 0; i < 6; i++)
            {
                if (Weights[i] > max)
                {
                    max = Weights[i];
                    if (i < 3)
                    {
                        Cl = i + 1;
                    }
                    else
                    {
                        Cl = i + 2;
                    }
                }
            }
            textBox13.Text = Cl.ToString();
        }
    }
}
