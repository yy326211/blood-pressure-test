using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ZedGraph;

namespace ppg_ecg
{
    public partial class Form1 : Form
    {
        [DllImport("CHIDClass.dll")]
        public static extern bool IniDevice(UInt16 mPID, UInt16 mVID);
        [DllImport("CHIDClass.dll")]
        public static extern bool DetectDevice();
        [DllImport("CHIDClass.dll")]
        public static extern bool CheckDeviceExist();
        [DllImport("CHIDClass.dll")]
        public static extern void GetAnalog(int mode, int ch, int range, double[] _readBuffer);
        [DllImport("CHIDClass.dll")]
        public static extern bool OutputAnalog(int ch, double val, int openClose);
        UInt16 PID = 0x1710;
        UInt16 VID = 0x5351;
        public Form1()
        {
            InitializeComponent();
        }
        const int number = 1024;
        double Ps, Pd;//舒张压，收缩压
        double rate;//脉率
        double K;//脉搏波特征值K
        double PVR;//外周阻力
        double AC;//顺应性
        double Td;//舒张期时间
        double Ts;
        double sv;       
        double[] ppg_data = new double[number];//采集的ppg数据
        double[] ecg_data = new double[number];//采集的ecg数据
        double[] ppg_datac = new double[number];
        int count0 = 0;
        int count1 = 0;
        int ch0 = 0;
        int ch1 = 1;
        int[] troughid = new int[40];
        int[] peakid = new int[40];
        double[] vol0 = new double[32];
        double[] vol1 = new double[32];
        double data0;
        double data1;        
        private void button1_Click(object sender, EventArgs e)
        {
            foundDev();
        }
        private bool foundDev()
        {
            bool result = false;

            IniDevice(PID, VID);
            if (DetectDevice())
            {
                result = true;
                textBox1.Text = "USB设备初始化正确";
            }
            else
            {
                textBox1.Text = "没有找到USB设备";
            }

            return result;
        }

        private void button2_Click(object sender, EventArgs e)
        {

            if (CheckDeviceExist())//检查USB是否有效，这个检查很重要
            {
                count0 = 0;
                count1 = 0;
                for (int k = 0; k < (number/8); k++)
                {
                    GetAnalog(0, ch0, 0, vol0);
                    for (int i = 0; i < 32; i = i + 4)
                    {
                        data0 = vol0[i];
                        ppg_data[count0] = data0;
                        count0++;
                    }
                }
                for (int k = 0; k < (number/8); k++)
                {
                    GetAnalog(0, ch1, 0, vol1);
                    for (int i = 0; i < 32; i = i + 4)
                    {
                        data1 = vol1[i];
                        ecg_data[count1] = data1+1;
                        count1++;
                    }
                }
            }          
            if (count0 == number && count1 == number)
            {
                MessageBox.Show("采集完毕！");
            }
            else
            {
                MessageBox.Show("采集失败，请重新采集！");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            graphshow();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            double sum0 = 0;
            double sum1 = 0;
            for (int i = 0; i < number - 7; i++)//八点平均滤波；
            {
                for (int j = 0; j < 8; j++)
                    sum0 += ppg_data[i + j];
                ppg_data[i] = sum0 / 8.0;
                sum0 = 0;
            }
            for (int i = 0; i < number - 7; i++)//八点平均滤波；
            {
                for (int j = 0; j < 8; j++)
                    sum1+= ecg_data[i + j];
                ecg_data[i] = sum1/8.0;
                sum1 = 0;
            }
            graphshow();
        }        
        private void button5_Click(object sender, EventArgs e)
        {
            delpiao(ppg_data, ppg_data, number);
            graphshow();
        }
        public void graphshow()//图像显示函数；
        {
            zgc1.GraphPane.CurveList.Clear();
            //get a reference to the GraphPane
            GraphPane myPane0 = zgc1.GraphPane;
            // Set the Titles
            myPane0.Title.Text = "脉搏波波型";
            myPane0.XAxis.Title.Text = "时间";
            myPane0.YAxis.Title.Text = "幅值";
            double x0;
            double y0;
            PointPairList list0 = new PointPairList();
            myPane0.XAxis.CrossAuto = true;
            zgc1.GraphPane.XAxis.Scale.MaxAuto = true;
            for (int j = 0; j < number; j++)
            {
                y0 = ppg_data[j];
                x0 = j;
                list0.Add(x0, y0);
            }
            LineItem myCurve = myPane0.AddCurve("", list0, Color.Red, SymbolType.None);
            zgc1.AxisChange();
            zgc1.Refresh();
            zgc2.GraphPane.CurveList.Clear();
            //get a reference to the GraphPane
            GraphPane myPane1 = zgc2.GraphPane;
            // Set the Titles
            myPane1.Title.Text = "ECG波型";
            myPane1.XAxis.Title.Text = "时间";
            myPane1.YAxis.Title.Text = "幅值";
            double x1;
            double y1;
            PointPairList list1 = new PointPairList();
            myPane1.XAxis.CrossAuto = true;
            zgc2.GraphPane.XAxis.Scale.MaxAuto = true;
            for (int j = 0; j < number; j++)
            {
                y1 = ecg_data[j];
                x1 = j;
                list1.Add(x1, y1);
            }
            LineItem myCurve1 = myPane1.AddCurve("", list1, Color.Green, SymbolType.None);
            zgc2.AxisChange();
            zgc2.Refresh();
        }
        public void delpiao(double[] peak, double[] trough, int id)//去除漂移函数
        {
            int j = 0, m = 0;
            double[] troughavg = new double[300];
            double[] troughsum = new double[2400];
            for (int i = 50; i <= id - 51; i++)
            {
                int n = 0;
                for (int k = 1; k <= 50; k++)
                {
                    if (peak[i] > peak[i + k] && peak[i] >= peak[i - k])
                        n++;                 
                }
                if (n == 49)
                {
                    peakid[j] = i;
                    j++;
                }
            }
            for (int i = 50; i <= id - 51; i++)
            {
                int n = 0;
                for (int k = 1; k <= 50; k++)
                {
                    if (trough[i] < trough[i + k] && trough[i] <= trough[i - k])
                        n++;                   
                }
                if (n == 49)
                {
                    troughid[m] = i;
                    m++;
                }
            }
            ppg_datac = ppg_data;//脉搏波波型整体平均
            for (int i1 = troughid[1]; i1 < troughid[2]; i1++)
            {
                int idc;
                idc = i1 - troughid[1];
                for (int i2 = 1; i2 < m - 1; i2++)
                {
                    troughsum[idc] += ppg_data[troughid[i2] + idc];
                }
                ppg_datac[troughid[1] + idc] = troughsum[idc] / Convert.ToDouble(m - 2);
                for (int i3 = 1; i3 < m - 1; i3++)
                {
                    ppg_data[troughid[i3] + idc] = ppg_datac[troughid[1] + idc];
                }
            }
            double sum2 = 0;
            for (int i4 = 0; i4 < number - 7; i4++)//八点平均滤波；
            {
                for (int j1 = 0; j1 < 8; j1++)
                    sum2 += ppg_data[i4 + j1];
                ppg_data[i4] = sum2 / 8.0;
                sum2 = 0;
            }
        }

        private void button6_Click(object sender, EventArgs e)//数据分析
        {
            rate = heartrate(ppg_data, number);
            K = find_trough(ppg_data, ppg_data, number);
            Td= find_td(ppg_data, ppg_data, number, rate);
            Ts = 60.0 / rate - Td;
            if (Ts < 0.13||rate<60)
                Ts = Ts + 0.02;
            else if (Ts > 0.17||rate>90)
                Ts = Ts - 0.03;
            Ps = (210 - Ts * 650);//使用ts代替ptt
            sv = 0.283 * 60 / rate / (0.33 * 0.33) * 35.0;
            PVR = (Ps - 40.0 + 0.33 * 40.0) / sv * (60.0 / rate);
            AC = 0.2 * (60 / rate) / (0.33 * 0.33);
            if (K > 0.35 && K <= 0.39)//特征值校正
                Ps = Ps + 4;
            else if (K > 0.39 && K <= 0.45)
                Ps = Ps + 8;
            else if (K > 0.45)
                Ps = Ps + 12;
            Pd = Ps * Math.Exp(-Td / (PVR * AC));
            rate = Convert.ToInt32(rate);
            textBox2.Text = rate.ToString();                        
            textBox3.Text = Pd.ToString("0");
            textBox4.Text = Ps.ToString("0");
            textBox5.Text = K.ToString("0.00");
            textBox6.Text = Ts.ToString("0.0000");

        }
        public double heartrate(double[] peak, int id)
        {            
            int j = 0;
            double chazhi = 0, chazhihe = 0;
            double pulserate = 0;
            int[] peakid = new int[15];
            for (int i = 50; i <= id - 51; i++)
            {
                int n = 0;
                for (int k = 1; k < 50; k++)
                {
                    if (peak[i] > peak[i + k] && peak[i] >= peak[i - k])
                        n++;
                }
                if (n == 49)
                {
                    peakid[j] = i;
                    j++;

                }
            }
         for (int i = 2; i != j-2; i++)
        {
            chazhi = peakid[i + 1] - peakid[i];
            chazhihe += chazhi;
        }
            pulserate = 60 / ((chazhihe / (j-4)) /130.0);//采样频率
            return pulserate;
        }
        public double find_trough(double[] peak, double[] trough, int id)//实时特征值k的函数
        {
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();
            int j = 0, m = 0;
            double k1 = 0, k2 = 0, k3 = 0, k4 = 0, vg = 0, vm = 0;
            int[] troughid = new int[15];
            int[] peakid = new int[15];
            for (int i = 50; i <= id - 51; i++)
            {
                int n = 0;
                for (int k = 1; k <= 50; k++)
                {
                    if (peak[i] > peak[i + k] && peak[i] >= peak[i - k])
                        n++;
                }
                if (n == 49)
                {
                    peakid[j] = i;
                    j++;
                    comboBox1.Items.Add("第" + j + "个峰值号" + i);
                }

            }
            for (int i = 50; i <= id - 51; i++)
            {
                int n = 0;
                for (int k = 1; k <= 50; k++)
                {
                    if (trough[i] <trough[i + k] && trough[i] <= trough[i - k])
                        n++;
                }
                if (n == 50)
                {
                    troughid[m] = i;
                    m++;
                    comboBox2.Items.Add("第" + m + "个谷值号" + i);
                }
            }
            if (troughid[0] < peakid[0])
            {
                for (int i = troughid[3]; i < troughid[4]; i++)
                    vg = vg + trough[i];
                vm = vg / (troughid[4] - troughid[3]);
                k1 = (vm - trough[troughid[3]]) / (peak[peakid[3]] - peak[troughid[3]]);
                vg = 0;
                for (int i = troughid[4]; i < troughid[5]; i++)
                    vg = vg + trough[i];
                vm = vg / (troughid[5] - troughid[4]);
                k2 = (vm - trough[troughid[4]]) / (peak[peakid[4]] - peak[troughid[4]]);
                vg = 0;
                for (int i = troughid[5]; i < troughid[6]; i++)
                    vg = vg + trough[i];
                vm = vg / (troughid[6] - troughid[5]);
                k3 = (vm - trough[troughid[5]]) / (peak[peakid[5]] - peak[troughid[5]]);
                vg = 0;
                for (int i = troughid[6]; i < troughid[7]; i++)
                    vg = vg + trough[i];
                vm = vg / (troughid[7] - troughid[6]);
                k4 = (vm - trough[troughid[6]]) / (peak[peakid[6]] - peak[troughid[6]]);
                vg = 0;
            }
            else
            {
                for (int i = troughid[3]; i < troughid[4]; i++)
                    vg = vg + trough[i];
                vm = vg / (troughid[4] - troughid[3]);
                k1 = (vm - trough[troughid[3]]) / (peak[peakid[4]] - peak[troughid[3]]);
                vg = 0;
                for (int i = troughid[4]; i < troughid[5]; i++)
                    vg = vg + trough[i];
                vm = vg / (troughid[5] - troughid[4]);
                k2 = (vm - trough[troughid[4]]) / (peak[peakid[5]] - peak[troughid[4]]);
                vg = 0;
                for (int i = troughid[5]; i < troughid[6]; i++)
                    vg = vg + trough[i];
                vm = vg / (troughid[6] - troughid[5]);
                k3 = (vm - trough[troughid[5]]) / (peak[peakid[6]] - peak[troughid[5]]);
                vg = 0;
                for (int i = troughid[6]; i < troughid[7]; i++)
                    vg = vg + trough[i];
                vm = vg / (troughid[7] - troughid[6]);
                k4 = (vm - trough[troughid[6]]) / (peak[peakid[7]] - peak[troughid[6]]);
                vg = 0;
            }

            return (k1 + k2 + k3 + k4) / 4.0;
        }
        public double find_td(double[] peak, double[] trough, int id, double rate)//实时舒张期时间函数
        {
            int j = 0, m = 0;
            double td = 0, t = 0, td1 = 0, t1 = 0, td2 = 0, td3 = 0, td4 = 0, t2 = 0, t3 = 0, t4 = 0;
            int[] troughid = new int[15];
            int[] peakid = new int[15];
            for (int i = 50; i <= id - 51; i++)
            {
                int n = 0;
                for (int k = 1; k <= 50; k++)
                {
                    if (peak[i] > peak[i + k] && peak[i] >= peak[i - k])
                        n++;
                }
                if (n == 49)
                {
                    peakid[j] = i;
                    j++;
                }
            }
            for (int i = 50; i <= id - 51; i++)
            {
                int n = 0;
                for (int k = 1; k <= 50; k++)
                {
                    if (trough[i] < trough[i + k] && trough[i] <= trough[i - k])
                        n++;
                }
                if (n == 49)
                {
                    troughid[m] = i;
                    m++;
                }
            }
            if (troughid[0] < peakid[0])
            {
                td1 = troughid[3] - peakid[2];
                t1 = troughid[3] - troughid[2];
                td2 = troughid[4] - peakid[3];
                t2 = troughid[4] - troughid[3];
                td3 = troughid[5] - peakid[4];
                t3 = troughid[5] - troughid[4];
                td4 = troughid[6] - peakid[5];
                t4 = troughid[6] - troughid[5];
                t = (t1 + t2 + t3 + t4) / 4.0;
                td = (td1 + td2 + td3 + td4) / 4.0;
            }
            else
            {
                td1 = troughid[2] - peakid[2];
                t1 = troughid[2] - troughid[1];
                td2 = troughid[3] - peakid[3];
                t2 = troughid[3] - troughid[2];
                td3 = troughid[4] - peakid[4];
                t3 = troughid[4] - troughid[3];
                td4 = troughid[5] - peakid[5];
                t4 = troughid[5] - troughid[4];
                t = (t1 + t2 + t3 + t4) / 4.0;
                td = (td1 + td2 + td3 + td4) / 4.0;
            }
            return td/130.0;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            //保存到本地的路径
            string mFileFullname1 = @"C:\Users\yangyang\Desktop\ppgdata.txt";
            string mFileFullname2 = @"C:\Users\yangyang\Desktop\ecgdata.txt";
            //编写器
            System.IO.StreamWriter mStreamWriter1 = new System.IO.StreamWriter(mFileFullname1, false, System.Text.Encoding.UTF8);
            System.IO.StreamWriter mStreamWriter2 = new System.IO.StreamWriter(mFileFullname2, false, System.Text.Encoding.UTF8);
            double[] mStrs1 = new double[number];
            double[] mStrs2 = new double[number];
            mStrs1 = ppg_data;
            mStrs2 = ecg_data;
            for (int i = 0; i < mStrs1.Length; i++)
            {
                mStreamWriter1.WriteLine(mStrs1[i]);
            }
            for (int i = 0; i < mStrs2.Length; i++)
            {
                mStreamWriter2.WriteLine(mStrs2[i]);
            }
            //用完StreamWriter的对象后一定要及时销毁
            mStreamWriter1.Close();
            mStreamWriter1.Dispose();
            mStreamWriter1 = null;
            mStreamWriter2.Close();
            mStreamWriter2.Dispose();
            mStreamWriter2 = null;
        }

    }
}
