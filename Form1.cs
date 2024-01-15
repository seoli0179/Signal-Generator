using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Security;

namespace Serial2
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            #region 초기 설정
            Control_initSetting();
            #endregion
        }

        #region Control 동작

        public void Control_initSetting()
        {
            //주파수 단위: MHz
            comboBox1.SelectedIndex = 1;
            //파워 단위: dBm
            comboBox3.SelectedIndex = 0;
            //band rate : 9600
            comboBox4.SelectedIndex = 2;
            //Frequcey : 4000
            numericUpDown1.Maximum = 6000;
            numericUpDown1.DecimalPlaces = 3;
            numericUpDown1.Value = 4000;
            //Frequcey Incr : 100
            numericUpDown2.Value = 100;
            numericUpDown2.DecimalPlaces = 3;
            //Amplitude : -135
            numericUpDown3.Minimum = -1000;
            numericUpDown3.DecimalPlaces = 2;
            numericUpDown3.Value = -135;
            //Amplitude Incr : 5
            numericUpDown4.DecimalPlaces = 2;
            numericUpDown4.Value = 5;
            //Amplitude Limit
            numericUpDown5.Maximum = 5000;
            numericUpDown5.Value = 150;
            numericUpDown4.Maximum = numericUpDown5.Value;

            numericUpDown6.Value = 0;
            numericUpDown6.Minimum = -150;
            numericUpDown6.DecimalPlaces = 2;

            listView1.View = View.Details;
            listView1.GridLines = true;
            listView1.FullRowSelect = true;
            listView1.MultiSelect = false;
            

            listView1.LabelEdit = false;

            listView1.Columns.Add("번호", 40);
            listView1.Columns[0].TextAlign = HorizontalAlignment.Center;
            listView1.Columns.Add("주파수", 65);
            listView1.Columns[1].TextAlign = HorizontalAlignment.Center;
            listView1.Columns.Add("단위", 50);
            listView1.Columns[2].TextAlign = HorizontalAlignment.Center;
            listView1.Columns.Add("진폭", 40);
            listView1.Columns[3].TextAlign = HorizontalAlignment.Center;
            listView1.Columns.Add("단위", 50);
            listView1.Columns[4].TextAlign = HorizontalAlignment.Center;

            saveFileDialog1.DefaultExt = ".csv"; // Default file extension
            saveFileDialog1.Filter = "Save Room File (.csv)|*.csv";

            checkBox1.Checked = false;
            checkBox2.Checked = true;
            checkBox3.Checked = false;
            checkBox4.Checked = true;

            serialPort1.ReadTimeout = 500;

            comboBox5.SelectedIndex = 0;

        }

        #region listview 동작

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                Control_ListviewSubItemAdd(listView1.FocusedItem.Index + 1, Convert.ToDouble(numericUpDown1.Value), comboBox1.Text, Convert.ToDouble(numericUpDown3.Value), comboBox3.Text);
            }
            catch
            {
                Control_ListviewSubItemAdd(1, Convert.ToDouble(numericUpDown1.Value), comboBox1.Text, Convert.ToDouble(numericUpDown3.Value), comboBox3.Text);
            }
            

        }

        private void numericUpDown3_KeyDown(object sender, KeyEventArgs e)
        {

            if (e.KeyCode == Keys.Enter)
            {

                if (serialPort1.IsOpen)
                {
                    Control_LogWrite(Serial_AmplitudeControl(Convert.ToDouble(numericUpDown3.Value), comboBox3.Text));
                }
                else
                {
                    Control_LogWrite("Port Close!");
                }

            }

        }

        public void Control_ListviewClear()
        {
            listView1.Items.Clear();
        }

        public void Control_ListviewSubItemAdd(int indexNum, Double Frequecy, String F_Unit, Double Amplitude, String A_Unit)
        {
            try
            {
                listView1.Items.Insert(listView1.FocusedItem.Index + 1, new ListViewItem(new String[] { indexNum + "", Frequecy + "", F_Unit, Amplitude + "", A_Unit }));
                listView1.Items[listView1.FocusedItem.Index + 1].Focused = true;
            }
            catch
            {
                listView1.Items.Add(new ListViewItem(new String[] { 1 + "", Frequecy + "", F_Unit, Amplitude + "", A_Unit }));
                listView1.Items[0].Focused = true;

            }
            listView1.Items[listView1.FocusedItem.Index].Selected = true;

            for(int i=0; i<listView1.Items.Count;i++)
            {
                listView1.Items[i].SubItems[0].Text = Convert.ToString(i + 1);
            }

            //Control_LogWrite(listView1.FocusedItem + "");
        }

        ListViewItem.ListViewSubItem SelectedLSI;

        private void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            ListViewHitTestInfo i = listView1.HitTest(e.X, e.Y);
            SelectedLSI = i.SubItem;
            if (SelectedLSI == null)
                return;

            int border = 0;
            switch (listView1.BorderStyle)
            {
                case BorderStyle.FixedSingle:
                    border = 1;
                    break;
                case BorderStyle.Fixed3D:
                    border = 2;
                    break;
            }

            int CellWidth = SelectedLSI.Bounds.Width;
            int CellHeight = SelectedLSI.Bounds.Height;
            int CellLeft = border + listView1.Left + i.SubItem.Bounds.Left;
            int CellTop = listView1.Top + i.SubItem.Bounds.Top;
            // First Column
            if (i.SubItem == i.Item.SubItems[0])
                CellWidth = listView1.Columns[0].Width;

            TxtEdit.Location = new Point(CellLeft, CellTop);
            TxtEdit.Size = new Size(CellWidth, CellHeight);
            TxtEdit.Visible = true;
            TxtEdit.BringToFront();
            TxtEdit.Text = i.SubItem.Text;
            TxtEdit.Select();
            TxtEdit.SelectAll();
        }
        private void listView2_MouseDown(object sender, MouseEventArgs e)
        {
            HideTextEditor();
        }
        private void listView2_Scroll(object sender, EventArgs e)
        {
            HideTextEditor();
        }
        private void TxtEdit_Leave(object sender, EventArgs e)
        {
            HideTextEditor();
        }
        private void TxtEdit_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Return)
                HideTextEditor();
        }
        private void HideTextEditor()
        {
            TxtEdit.Visible = false;
            if (SelectedLSI != null)
                SelectedLSI.Text = TxtEdit.Text;
            SelectedLSI = null;
            TxtEdit.Text = "";
        }

        private void TxtEdit_FocusLeave(object sender, EventArgs e)
        {
            HideTextEditor();
        }

        private void TxtEdit_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                HideTextEditor();
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {

            try
            {
                listView1.Items[listView1.FocusedItem.Index].Remove();
                listView1.Items[listView1.FocusedItem.Index].Selected = true;
            }
            catch
            {
                Control_LogWrite("listview1 subitem이 전부 제거됨");
            }

        }

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                listView1.Items[listView1.FocusedItem.Index - 1].Focused = true;
                listView1.Items[listView1.FocusedItem.Index].Selected = true;
                if(Control_RoomSet(Convert.ToDouble( listView1.Items[listView1.FocusedItem.Index].SubItems[1].Text), listView1.Items[listView1.FocusedItem.Index].SubItems[2].Text, Convert.ToDouble(listView1.Items[listView1.FocusedItem.Index].SubItems[3].Text), listView1.Items[listView1.FocusedItem.Index].SubItems[4].Text))
                {
                    Control_LogWrite("Room "+(listView1.FocusedItem.Index + 1) + "번 설정 완료");
                }
                else
                {
                    Control_LogWrite(listView1.FocusedItem.Index + ": 오류 발생");
                }
            }
            catch
            {
                listView1.Items[listView1.FocusedItem.Index].Focused = true;
                listView1.Items[listView1.FocusedItem.Index].Selected = true;
                if (Control_RoomSet(Convert.ToDouble(listView1.Items[listView1.FocusedItem.Index].SubItems[1].Text), listView1.Items[listView1.FocusedItem.Index].SubItems[2].Text, Convert.ToDouble(listView1.Items[listView1.FocusedItem.Index].SubItems[3].Text), listView1.Items[listView1.FocusedItem.Index].SubItems[4].Text))
                {
                    Control_LogWrite("Room " + (listView1.FocusedItem.Index + 1) + "번 설정 완료");
                }
                else
                {
                    Control_LogWrite(listView1.FocusedItem.Index + ": 오류 발생");
                }
                Control_LogWrite("listview1 index가 최상위임");
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                listView1.Items[listView1.FocusedItem.Index + 1].Focused = true;
                listView1.Items[listView1.FocusedItem.Index].Selected = true;
                if (Control_RoomSet(Convert.ToDouble(listView1.Items[listView1.FocusedItem.Index].SubItems[1].Text), listView1.Items[listView1.FocusedItem.Index].SubItems[2].Text, Convert.ToDouble(listView1.Items[listView1.FocusedItem.Index].SubItems[3].Text), listView1.Items[listView1.FocusedItem.Index].SubItems[4].Text))
                {
                    Control_LogWrite(listView1.FocusedItem.Index + ": Setting 완료");
                }
                else
                {
                    Control_LogWrite(listView1.FocusedItem.Index + ": 오류 발생");
                }
            }
            catch
            {
                listView1.Items[listView1.FocusedItem.Index].Focused = true;
                listView1.Items[listView1.FocusedItem.Index].Selected = true;
                if (Control_RoomSet(Convert.ToDouble(listView1.Items[listView1.FocusedItem.Index].SubItems[1].Text), listView1.Items[listView1.FocusedItem.Index].SubItems[2].Text, Convert.ToDouble(listView1.Items[listView1.FocusedItem.Index].SubItems[3].Text), listView1.Items[listView1.FocusedItem.Index].SubItems[4].Text))
                {
                    Control_LogWrite(listView1.FocusedItem.Index + ": Setting 완료");
                }
                else
                {
                    Control_LogWrite(listView1.FocusedItem.Index + ": 오류 발생");
                }
                Control_LogWrite("listview1 index가 최하위임");
            }
        }

        #endregion

        public Boolean Control_RoomSet(Double Frequecy, String F_Unit, Double Amplitude, String A_Unit)
        {
            try
            {
                Control_LogWrite(Serial_FrequecyControl(Frequecy, F_Unit));
                Control_LogWrite(Serial_AmplitudeControl(Amplitude, A_Unit));
                return true;
            }
            catch
            {
                return false;
            }
        }


        //numericUpDown1 Increment을 numericUpDown2 Value로 결정
        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            numericUpDown1.Increment = numericUpDown2.Value;
        }

        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {
            numericUpDown3.Increment = numericUpDown4.Value;
        }

        public void Control_LogWrite(List<String> str)
        {
            for (int i = 0; i < str.Count; i++)
            {
                textBox1.AppendText(Environment.NewLine + Environment.NewLine + str[i]);
            }
        }

        public void Control_LogWrite(String str)
        {
            textBox1.AppendText(Environment.NewLine + Environment.NewLine + str);
        }

        private void numericUpDown5_ValueChanged(object sender, EventArgs e)
        {
            numericUpDown3.Maximum = numericUpDown5.Value;
        }

        private void numericUpDown1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {

                if (serialPort1.IsOpen)
                {
                    Control_LogWrite(Serial_FrequecyControl(Convert.ToDouble(numericUpDown1.Value), comboBox1.Text));
                }

            }
        }

        #endregion

        #region Serial 컨트롤
        public String[] Serial_serchPortList()
        {
            return System.IO.Ports.SerialPort.GetPortNames();
        }

        public List<String> Serial_PortOpen(String PortName, int baudrate)
        {
            List<String> str = new List<string>();
            try
            {
                if (PortName == "" || PortName == null || baudrate == 0)
                {
                    str.Add("포트 세팅이 안 되어 있습니다.");        //사용가능한 포트, BaudRate 지정안할시 처리
                }
                else
                {
                    if (serialPort1.IsOpen)
                    {
                        str.Add(serialPort1.PortName + " 에 이미 연결되어 있습니다.");
                    }
                    else
                    {
                        serialPort1.PortName = PortName;
                        serialPort1.BaudRate = baudrate;//콤보박스에 있는 데이터는 문자이기때문에 Int로 형변환
                        serialPort1.DataBits = 8;   //기본 데이터 비트 지정
                        serialPort1.Open();                     //선택한 serialPort 오픈
                        str.Add(PortName + " / " + baudrate + " 연결 완료");
                    }

                }
            }
            catch (UnauthorizedAccessException)      //접근 불가 예외처리
            {
                str.Add("접근 불가");
            }

            return str;
        }

        public String Serial_PortClose()
        {
            String str = serialPort1.PortName;
            serialPort1.Close();
            return str + " 연결 해제";
        }

        public List<String> Serial_WriteLine(String text)
        {
            List<String> str = new List<string>();
            try
            {
                if (serialPort1.IsOpen)
                {
                    serialPort1.WriteLine(text);
                    str.Add("Tx: " + text);
                }
                else
                {
                    str.Add("포트가 연결되어 있지 않음");
                }
            }
            catch (Exception e)
            {
                str.Add(e.Message);
            }
            return str;
        }

        public List<String> Serial_FrequecyControl(Double Frequecy, String Unit)
        {
            return Serial_WriteLine(":frequency " + Frequecy + Unit);
        }

        public List<String> Serial_OffsestControl(Double Offsest, String Unit)
        {
            return Serial_WriteLine(":power:offset " + Offsest + " " + Unit);
        }

        public List<String> Serial_AmplitudeControl(Double Amplitude, String Unit)
        {
            return Serial_WriteLine(":power " + Amplitude + Unit);
        }

        public List<String> Serial_PowerUnitControl(String Unit)
        {
            List<String> str = new List<string>();
            if (Unit == "dBm")
            {
                return Serial_WriteLine(":UNIT:POWer dBm");
            }
            if (Unit == "uV")
            {
                return Serial_WriteLine(":UNIT:POWer V");
            }
            str.Add("ERROR");
            return str;

        }

        public List<String> Serial_PowerControl(String Power)
        {
            return Serial_WriteLine(":output " + Power);
        }

        public List<String> Serial_ModControl(String Mod)
        {
            return Serial_WriteLine(":output:modulation " + Mod);
        }

        public String Serial_WriteLineRead(String text)
        {
            Control_LogWrite(Serial_WriteLine(text));
            return "Rx: "+serialPort1.ReadLine();
        }

        public Double Serial_PowerRead(String Unit)
        {
            List<String> str = new List<string>();

            Double value = 0;

            Control_LogWrite(Serial_PowerUnitControl(Unit));
            Control_LogWrite(Serial_WriteLine(":power?"));
            if (Unit == "dBm")
            {
                value = Convert.ToDouble(serialPort1.ReadLine());
            }
            if (Unit == "uV")
            {
                value = Convert.ToDouble(serialPort1.ReadLine()) * 1000000;
            }

            str.Add("Read Amplitude : " + value + " " + Unit);

            Control_LogWrite(str);

            return value;
        }

        public Double Serial_FrequecyRead(String Unit)
        {
            List<String> str = new List<string>();

            Double value = 0;

            Control_LogWrite(Serial_WriteLine(":frequency?"));

            if (Unit == "GHz")
            {
                value = Convert.ToDouble(serialPort1.ReadLine()) / 1000000000;
            }
            else if (Unit == "MHz")
            {
                value = Convert.ToDouble(serialPort1.ReadLine()) / 1000000;
            }
            else
            {
                value = Convert.ToDouble(serialPort1.ReadLine()) / 1000;
            }

            str.Add("Rx : " + value);

            Control_LogWrite(str);

            return value;
        }

        public Double Serial_AmplitudeOffsetRead()
        {
            List<String> str = new List<string>();

            Double value = 0;

            Control_LogWrite(Serial_WriteLine(":power:offset?"));
            value = Convert.ToDouble(serialPort1.ReadLine());

            str.Add("Rx : " + value);

            Control_LogWrite(str);

            return value;
        }

        public List<String> Serial_UnitRead()
        {
            List<String> str = new List<string>();
            Serial_WriteLine(":UNIT:POWer?");
            str.Add("Rx: "+serialPort1.ReadLine());
            return str;
        }


        #endregion

        #region Serial 동작

        private void comboBox2_Click(object sender, EventArgs e)
        {
            comboBox2.Items.Clear();
            String[] portList = Serial_serchPortList();
            for (int i = 0; i < portList.Length; i++)
            {
                comboBox2.Items.Add(portList[i]);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                Control_LogWrite(Serial_PortOpen(comboBox2.Text, Convert.ToInt32(comboBox4.Text)));

                Control_LogWrite(Serial_WriteLineRead("*IDN?"));

                Control_LogWrite(Serial_PowerUnitControl(comboBox3.Text));

                Control_LogWrite(Serial_PowerControl("OFF"));

                Control_LogWrite(Serial_ModControl("OFF"));

                checkBox1.Checked = false;
                checkBox2.Checked = true;
                checkBox3.Checked = false;
                checkBox4.Checked = true;

                numericUpDown1.Value = Convert.ToDecimal(Serial_FrequecyRead(comboBox1.Text));
            }
            catch (Exception ex)
            {
                Control_LogWrite(ex.Message);
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            Control_LogWrite(Serial_PortClose());
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                Control_LogWrite(Serial_FrequecyControl(Convert.ToDouble(numericUpDown1.Value), comboBox1.Text));
            }

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                Control_LogWrite(Serial_FrequecyControl(Convert.ToDouble(numericUpDown1.Value), comboBox1.Text));
            }

        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                Control_LogWrite(Serial_AmplitudeControl(Convert.ToDouble(numericUpDown3.Value), comboBox3.Text));
            }

        }
        private void button7_Click(object sender, EventArgs e)
        {
            Control_LogWrite(Serial_PowerControl("On"));
            checkBox1.Checked = true;
            checkBox2.Checked = false;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Control_LogWrite(Serial_PowerControl("Off"));
            checkBox1.Checked = false;
            checkBox2.Checked = true;
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                try
                {
                    if (comboBox3.Text == "dBm")
                    {
                        numericUpDown3.Value = -135;
                    }
                    if (comboBox3.Text == "uV")
                    {
                        numericUpDown3.Value = 5;
                    }

                }
                catch (Exception ex)
                {
                    Control_LogWrite(ex.Message);
                }

            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Control_LogWrite(Serial_ModControl("On"));
            checkBox3.Checked = true;
            checkBox4.Checked = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Control_LogWrite(Serial_ModControl("Off"));
            checkBox3.Checked = false;
            checkBox4.Checked = true;
        }


        #endregion

        #region 파일 입출력
        private void button12_Click(object sender, EventArgs e)
        {
            
        }

        private void button11_Click(object sender, EventArgs e)
        {

           

        }

        #endregion

        #region 메뉴 동작

        private void 파일ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void 저장ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {

                    StreamWriter sw = new StreamWriter(saveFileDialog1.FileName, false);

                    Control_LogWrite("System: " + saveFileDialog1.FileName + " 저장하기");
                    //Control_ListviewClear();

                    for (int i = 0; i < listView1.Items.Count; i++)
                    {
                        Control_LogWrite(listView1.Items[i].SubItems[0].Text + "," + listView1.Items[i].SubItems[1].Text + "," + listView1.Items[i].SubItems[2].Text + "," + listView1.Items[i].SubItems[3].Text + "," + listView1.Items[i].SubItems[4].Text);
                        sw.WriteLine(listView1.Items[i].SubItems[0].Text + "," + listView1.Items[i].SubItems[1].Text + "," + listView1.Items[i].SubItems[2].Text + "," + listView1.Items[i].SubItems[3].Text + "," + listView1.Items[i].SubItems[4].Text);
                    }

                    sw.Close();
                    sw.Dispose();

                }
                catch (SecurityException ex)
                {
                    Control_LogWrite($"Security error.\n\nError message: {ex.Message}\n\n" + $"Details:\n\n{ex.StackTrace}");
                }
                catch (Exception exd)
                {
                    Control_LogWrite(exd.Message);
                }
            }
        }

        private void 불러오기ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {

                    var sr = new StreamReader(openFileDialog1.FileName);
                    Control_LogWrite("System: " + openFileDialog1.FileName + " 불러오기");
                    Control_ListviewClear();

                    while (!sr.EndOfStream)
                    {

                        // 한 줄씩 읽어온다.

                        String line = sr.ReadLine();

                        // 쉼표( , )를 기준으로 데이터를 분리한다.

                        String[] data = line.Split(',');

                        Control_ListviewSubItemAdd(Convert.ToInt32(data[0]), Convert.ToDouble(data[1]), data[2], Convert.ToDouble(data[3]), data[4]);

                        // 결과를 출력해본다.
                        String str = "";

                        for (int i = 0; i < data.Length; i++)
                        {
                            str += data[i] + ",";
                        }

                        Control_LogWrite(str);

                    }
                    //Control_LogWrite(sr.ReadToEnd());


                    sr.Close();
                    sr.Dispose();


                }
                catch (SecurityException ex)
                {
                    Control_LogWrite($"Security error.\n\nError message: {ex.Message}\n\n" + $"Details:\n\n{ex.StackTrace}");
                }
                catch (Exception exd)
                {
                    Control_LogWrite(exd.Message);
                }
            }
        }


        #endregion

        private void 진폭OffsetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if(serialPort1.IsOpen)
            {
                Form2 frm2 = new Form2(); // Form2형 frm2 인스턴스화(객체 생성)

                frm2.Passvalue = Serial_AmplitudeOffsetRead();  // 전달자(Passvalue)를 통해서 Form2 로 전달
                frm2.ShowDialog();
            }
            else
            {
                Control_LogWrite("Port Close!");
            }


        }

        private void numericUpDown6_ValueChanged(object sender, EventArgs e)
        {
            //Control_LogWrite("a");
            Control_LogWrite(Serial_OffsestControl(Convert.ToDouble(numericUpDown6.Value), comboBox5.Text));
 
        }

        private void numericUpDown6_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {

                if (serialPort1.IsOpen)
                {
                    Control_LogWrite(Serial_OffsestControl(Convert.ToDouble(numericUpDown6.Value), comboBox5.Text));
                }
                else
                {
                    Control_LogWrite("Port Close!");
                }

            }
        }
    }


}
