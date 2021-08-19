using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.IO;
using System.Management;
using System.Diagnostics;


namespace MW_001_CodeWriter
{
    public partial class Form1 : Form
    {
        StreamWriter LOG;
        StreamReader SRead;
        DateTime startTime;
        string filePath;
        string tellnum;
        string[,] portNames;
        string[] RxData;
        string dataIN;
        string FDTI1;
        string FDTI2;

        bool errFlag = false;
        bool endFlag = false;
        bool statusFlag = false;
        bool timeOut = false;
        bool startUp = false;
        bool idleFlag = false;
        bool cityFlag = false;
        bool sensFlag = false;

        public Form1()
        {
            InitializeComponent();

            //Time stamp
            DateTime start_time = DateTime.Now;
            string stime = start_time.ToString("yyyyMMdd");
            Console.WriteLine("LOG: " + stime);
            //Start logging
            LOG = new StreamWriter(stime + ".log", true, System.Text.Encoding.Default);
            LOG.WriteLine(Environment.NewLine + start_time);
            Console.WriteLine("LOG: Start logging.");
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            FDTI1 = "FTDIBUS\\VID_0403+PID_6001+00000000A\\0000";
            FDTI2 = "FTDIBUS\\VID_0403+PID_6001+FT4U1V7RA\\0000";
            toolStripProgressBar1.Value = 0;
            button_action.Text = "選択";
            button_action.Enabled = false;
            button_stop.Text = "停止";
            textBox_csv.Clear();
            textBox_devicecode.Clear();
            textBox_citycode.Clear();
            textBox_tellnumber.Clear();
            comboBox_comport.Items.Clear();
            panel2.Visible = false;
            Console.WriteLine("Form1_Load");
        }

        private void Form1_Activated(object sender, EventArgs e)
        {
            Console.WriteLine("LOG: Form1_Activated");
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            Application.DoEvents();
            Console.WriteLine("LOG: Form1_Shown");
            CSVSearch();
            PortSearch();
        }

        private void CSVSearch()
        {
            //File Path
            String appPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            appPath = appPath.Replace("MW-001_CodeWriter.exe", "");
            Console.WriteLine("LOG: " + appPath + Environment.NewLine);

            //CSV File Search
            string[] CSVfiles = System.IO.Directory.GetFiles(@appPath, "*.csv", System.IO.SearchOption.AllDirectories);
            if (CSVfiles.Length == 1) //There is a csv file.
            {
                filePath = CSVfiles[0];
                LOG.WriteLine(filePath);
                textBox_csv.Text = Path.GetFileName(CSVfiles[0]);
                toolStripProgressBar1.Value = 5; //action-0
                toolStripStatusLabel1.Text = "水位計IDファイル読込済";
            }
            else //None or many csv files.
            {
                toolStripStatusLabel1.Text = "水位計IDファイルエラー";
                errFlag = true;
            }

            LOG.WriteLine(toolStripStatusLabel1.Text);
            
        }
        private void PortSearch()
        {
            string[] ports;
            
            if (errFlag == true)
            {

                DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                if (result == DialogResult.OK)
                {
                    LOG.Close();
                    Application.Exit();
                }
            }
            else if (errFlag == false)
            {
                toolStripStatusLabel1.Text = "ケーブル検索中";

                do
                {
                    Application.DoEvents();
                    if (endFlag == true)
                    {
                        break;
                    }

                    ports = SerialPort.GetPortNames();

                    if (ports.Length > 0)
                    {
                        comboBox_comport.Items.Clear();
                        portNames = GetDeviceNames();
                        for (int i = 0; i < ports.Length; i++)
                        {
                            if (portNames[i,1].Contains(FDTI2))
                            {
                                comboBox_comport.Items.Add(portNames[i, 0]); //.Substring(0, portNames[i].IndexOf("(")));
                            }
                            
                        }
                        /*foreach (string str in portNames)
                        {
                            //if(str ==)
                            comboBox_comport.Items.Add(str);//.Substring(0, str.IndexOf("(")));
                            LOG.WriteLine("LIST:" + str);
                        }*/

                        int comIndex = comboBox_comport.FindString("USB Serial Port");
                        if (comIndex >= 0)
                        {
                            comboBox_comport.SelectedIndex = comIndex;
                        }

                        toolStripStatusLabel1.Text = "ケーブル選択";
                        toolStripProgressBar1.Value = 10;
                        button_action.Enabled = true;
                        idleFlag = true;
                    }
                } while (ports.Length == 0);
            }


            if (endFlag == true)
            {
                toolStripStatusLabel1.Text = "停止しますか";
                LOG.WriteLine(toolStripStatusLabel1.Text);
                DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "YesNo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    LOG.WriteLine("Yes");
                    LOG.Close();
                    Application.Exit();
                    return;
                }
                if (result == DialogResult.No)
                {
                    LOG.WriteLine("No");
                    endFlag = false;
                    PortSearch();
                }
            }
        }

        private void button_stop_Click(object sender, EventArgs e)
        {
            endFlag = true;
            toolStripStatusLabel1.Text = "停止ボタン";
            LOG.WriteLine(toolStripStatusLabel1.Text);

            if (idleFlag == true)
            {
                if(startUp == true)
                {
                    toolStripStatusLabel1.Text = "終了します。";
                    DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "終了", MessageBoxButtons.OK, MessageBoxIcon.None);

                    if (result == DialogResult.OK)
                    {
                        LOG.WriteLine("完了");
                        LOG.Close();
                        Application.Exit();
                    }
                    return;
                }
                else
                {
                    toolStripStatusLabel1.Text = "停止しますか";
                    LOG.WriteLine(toolStripStatusLabel1.Text);
                    DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "停止ボタン", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        LOG.WriteLine("Yes");
                        if (serialPort1.IsOpen)
                        {
                            serialPort1.Close();
                        }
                        LOG.Close();
                        Application.Exit();
                        return;
                    }
                    if (result == DialogResult.No)
                    {
                        LOG.WriteLine("No");
                        endFlag = false;
                    }
                    return;
                }
                

            }
        }
        private void button_action_Click(object sender, EventArgs e)
        {
            if (statusFlag == false && serialPort1.IsOpen == false)
            {
                button_action.Enabled = false;
                idleFlag = false;
                PortConnect();

                if (endFlag == true)
                {
                    LOG.Close();
                    Application.Exit();
                    return;
                }
                
                button_action.Text = "書込";
                NextPannel();

                toolStripStatusLabel1.Text = "所定の位置に磁石を近づけて、電源を入れて下さい。";
                PowerON();

                if (startUp == true)
                {
                    textBox_tellnumber.Text = tellnum;
                }

                if(statusFlag == true)
                {
                    button_action.Enabled = true;
                    idleFlag = true;
                    startUp = false;
                    toolStripStatusLabel1.Text = "水位計IDの書き込みが可能です。";
                }
                return;

            }
            if(statusFlag == true)
            {
                button_action.Enabled = false;
                idleFlag = false;
                
                WriteCommand("CITYCODE");
                CodeWrite();

                if(startUp == true)
                {
                    startUp = false;
                    WriteCommand("SENSORNO");
                    CodeWrite();
                }

                if(startUp == true)
                {
                    startUp = false;
                    WriteCommand("INFO");
                    CodeTest();
                }

                if(startUp == true)
                {
                    startUp = false;
                    WriteCommand("ATTACH");
                    AttachTest();
                }

                if(startUp == true)
                {
                    button_action.Text = "完了";
                    button_stop.Text = "終了";
                    idleFlag = true;
                    toolStripProgressBar1.Value = 100;
                    toolStripStatusLabel1.Text = "水位計ID書込み完了";
                    LOG.WriteLine(toolStripStatusLabel1.Text);
                }

            }
        }

        private void PortConnect()
        {
            try
            {
                toolStripStatusLabel1.Text = "接続中";
                this.Update();

                //serialport設定
                string strValue = portNames[comboBox_comport.SelectedIndex,0];
                strValue = strValue.Remove(0, strValue.IndexOf("(") + 1);
                strValue = strValue.Remove(strValue.IndexOf(")"));

                serialPort1.PortName = strValue;
                serialPort1.BaudRate = 115200; // Convert.ToInt32(cBoxBAUDRATE.Text);
                serialPort1.DataBits = 8; // Convert.ToInt32(cBoxDATABITS.Text);
                serialPort1.StopBits = (StopBits)Enum.Parse(typeof(StopBits), "One");
                serialPort1.Parity = (Parity)Enum.Parse(typeof(Parity), "None");
                //serialPort1.DataReceived += new SerialDataReceivedEventHandler(serialPort1_DataReceived);

                serialPort1.Open();
                System.Threading.Thread.Sleep(1000);
                if (serialPort1.IsOpen)
                {
                    toolStripStatusLabel1.Text = "接続済み";
                    LOG.WriteLine("SELECT:" + portNames[comboBox_comport.SelectedIndex,0]); //COM名
                    LOG.WriteLine(toolStripStatusLabel1.Text); //ケーブル選択済み

                    //次のステップを表示
                    toolStripProgressBar1.Value = 15;
                }
                else
                {
                    throw new Exception("Error");
                }
            }
            catch
            {
                toolStripStatusLabel1.Text = "ケーブル抜け";
                LOG.WriteLine(toolStripStatusLabel1.Text);
                DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                if (result == DialogResult.OK)
                {
                    //LOG.Close();
                    //Application.Exit();
                    endFlag = true;
                }
            }
        }

        private void NextPannel()
        {
            panel1.Visible = false;
            panel2.Visible = true;
        }

        private void PowerON()
        {
            if (serialPort1.IsOpen)
            {
                serialPort1.DiscardInBuffer();
                serialPort1.DiscardOutBuffer();
            }
            else
            {
                toolStripStatusLabel1.Text = "ケーブルが抜けました。";
                LOG.WriteLine(toolStripStatusLabel1.Text);
                DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                if (result == DialogResult.OK)
                {
                    LOG.Close();
                    Application.Exit();
                }
                return;
            }
            try
            {
                while (startUp == false && endFlag == false && errFlag == false)
                {
                    this.Activate();
                    this.Update();
                    Application.DoEvents();

                    if(timeOut == true)
                    {
                        DateTime endDT = DateTime.Now;
                        TimeSpan ts = endDT - startTime;
                        //Console.WriteLine(ts);
                        if (ts.TotalSeconds > 40)
                        {
                            //タイムアウトした
                            toolStripStatusLabel1.Text = "起動できません。(40秒タイムアウト)";
                            LOG.WriteLine(toolStripStatusLabel1.Text);
                            DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            if (result == DialogResult.OK)
                            {
                                serialPort1.Close();
                                LOG.Close();
                                Application.Exit();
                            }
                            return;
                        }
                    }
                    if (serialPort1.BytesToRead > 2)
                    {
                        dataIN += serialPort1.ReadExisting();
                        if(dataIN.IndexOf("\n") > 0)
                        {
                            RxData = dataIN.Split('\n'); //dataIN.Split(new string[] { "\n" }, StringSplitOptions.None); //Split('\n')
                            foreach (String s in RxData)
                            {
                                Console.WriteLine("LOG: " + s);
                            }
                            this.Invoke(new EventHandler(SerialLog));
                        }
                        dataIN = string.Empty;
                    }
                }
                if(startUp == true)
                {
                    toolStripStatusLabel1.Text = "起動完了";
                    Console.WriteLine(toolStripStatusLabel1.Text);
                    LOG.WriteLine(toolStripStatusLabel1.Text);
                    return;
                }
                if(endFlag == true)
                {
                    toolStripStatusLabel1.Text = "起動中停止";
                    LOG.WriteLine(toolStripStatusLabel1.Text);
                    DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (result == DialogResult.OK)
                    {
                        serialPort1.Close();
                        LOG.Close();
                        Application.Exit();
                    }
                    return;
                }
                if (errFlag == true)
                {
                    DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (result == DialogResult.OK)
                    {
                        serialPort1.Close();
                        LOG.Close();
                        Application.Exit();
                    }
                    return;
                }
            }
            catch
            {
                timeOut = false;
                toolStripStatusLabel1.Text = "起動できません。(その他)";
                Console.WriteLine(toolStripStatusLabel1.Text);
                LOG.WriteLine(toolStripStatusLabel1.Text);
                DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                if (result == DialogResult.OK)
                {
                    serialPort1.Close();
                    LOG.Close();
                    Application.Exit();
                }
            }
        }

        private void SerialLog(object sender, EventArgs e)
        {
            foreach (string s in RxData)
            {
                if (s.Contains("start"))
                {
                    toolStripStatusLabel1.Text = "磁石が反応してません。電源入れ直ししてください。";
                    Console.WriteLine(toolStripStatusLabel1.Text);
                    LOG.WriteLine(toolStripStatusLabel1.Text);
                    DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    if (result == DialogResult.OK)
                        return;
                }

                if (s.Contains("START TEST"))
                {
                    toolStripProgressBar1.Value = 20;
                    startTime = DateTime.Now; //時間取得
                    Console.WriteLine("LOG: " + startTime);
                    timeOut = true; //タイマー起動
                    toolStripStatusLabel1.Text = "制御部 起動開始 [起動完了まで35秒] [強制終了まで40秒]";
                    Console.WriteLine(toolStripStatusLabel1.Text);
                    LOG.WriteLine(toolStripStatusLabel1.Text); 
                    return;
                }

                if (s.Contains("WAKEUP"))
                {
                    toolStripProgressBar1.Value = 25;
                    toolStripStatusLabel1.Text = "通信部 起動開始 [起動完了まで20秒] [強制終了まで25秒]";
                    Console.WriteLine(toolStripStatusLabel1.Text);
                    LOG.WriteLine(toolStripStatusLabel1.Text); //テストモード起動中
                    return;
                }

                if (s.Contains("NUM=0"))
                {
                    timeOut = false;
                    int len = s.Length;
                    if (len < 15)
                    {
                        toolStripStatusLabel1.Text = "SIMカードを認識できません。";
                        Console.WriteLine(toolStripStatusLabel1.Text);
                        LOG.WriteLine(toolStripStatusLabel1.Text);
                        
                        errFlag = true;
                        return;
                    }
                    else
                    {
                        toolStripProgressBar1.Value = 30;
                        toolStripStatusLabel1.Text = "電話番号取得";
                        Console.WriteLine(toolStripStatusLabel1.Text);

                        //電話番号
                        tellnum = s.Substring(len - 11);
                        Console.WriteLine("LOG: " + tellnum);
                        LOG.WriteLine("電話番号: " + tellnum);

                        //次へ進む
                        startUp = true;
                    }
                }
            }
        }

        private void textBox_tellnumber_TextChanged(object sender, EventArgs e)
        {
            textBox_citycode.Clear();
            textBox_devicecode.Clear();

            if (textBox_tellnumber.Text.Length == 11) //11
            {
                try
                {
                    SRead = new StreamReader(@filePath, Encoding.Default);

                    string dat;
                    while ((dat = SRead.ReadLine()) != null)
                    {
                        string callnum;
                        string[] sbuf = dat.Split(',');
                        callnum = sbuf[0];

                        if (textBox_tellnumber.Text == callnum)
                        {
                            textBox_citycode.Text = sbuf[1];
                            textBox_devicecode.Text = sbuf[2];
                            toolStripProgressBar1.Value = 40;

                            //水位計ID書込み可
                            statusFlag = true;
                            break;
                        }
                        
                    }
                    SRead.Close();

                    if (statusFlag == false)
                    {
                        toolStripStatusLabel1.Text = "このSIMは登録がありません。";
                        LOG.WriteLine(toolStripStatusLabel1.Text);
                        DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                        if (result == DialogResult.OK)
                        {
                            serialPort1.Close();
                            LOG.Close();
                            Application.Exit();
                        }
                        SRead.Close();
                        return;
                    }
                    else
                    {
                        //log
                        LOG.WriteLine("市町村コード: " + textBox_citycode.Text);
                        LOG.WriteLine("水位計番号: " + textBox_devicecode.Text);
                    }
                }
                catch
                {
                    toolStripStatusLabel1.Text = "水位計ID検索時エラー";
                    LOG.WriteLine(toolStripStatusLabel1.Text);
                    DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    if (result == DialogResult.OK)
                    {
                        serialPort1.Close();
                        LOG.Close();
                        Application.Exit();
                    }
                     return;
                }
            }
            else
            {
                toolStripStatusLabel1.Text = "電話番号桁数エラー";
                LOG.WriteLine(toolStripStatusLabel1.Text);
                DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                if (result == DialogResult.OK)
                {
                    serialPort1.Close();
                    LOG.Close();
                    Application.Exit();
                }
                return;
            }
        }

        private void CodeWrite()
        {
            if (serialPort1.IsOpen)
            {
                serialPort1.DiscardInBuffer();
                serialPort1.DiscardOutBuffer();
            }
            else
            {
                toolStripStatusLabel1.Text = "ケーブルが抜けました。";
                LOG.WriteLine(toolStripStatusLabel1.Text);
                DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                if (result == DialogResult.OK)
                {
                    LOG.Close();
                    Application.Exit();
                }
                return;
            }
            try
            {
                while (startUp == false && endFlag == false)
                {
                    this.Activate();
                    this.Update();
                    Application.DoEvents();

                    if (timeOut == true)
                    {
                        DateTime endDT = DateTime.Now;
                        TimeSpan ts = endDT - startTime;
                        //Console.WriteLine(ts);
                        if (ts.TotalSeconds > 10)
                        {
                            //タイムアウトした
                            toolStripStatusLabel1.Text = "書き込みできません。(10秒タイムアウト)";
                            LOG.WriteLine(toolStripStatusLabel1.Text);
                            DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            if (result == DialogResult.OK)
                            {
                                serialPort1.Close();
                                LOG.Close();
                                Application.Exit();
                            }
                            return;
                        }
                    }
                    if (serialPort1.BytesToRead > 0)
                    {
                        
                        dataIN += serialPort1.ReadExisting();//ReadExisting();

                        if(dataIN.IndexOf("\n") > 0)
                        {
                            RxData = dataIN.Split('\n'); //dataIN.Replace("\r\n", "").Split('\n');
                            foreach (String s in RxData)
                            {
                                Console.WriteLine("LOG: " + s);
                            }
                            this.Invoke(new EventHandler(SerialWrite));
                        }
                        dataIN = string.Empty;
                    }
                }
                if (startUp == true)
                {
                    toolStripStatusLabel1.Text = "書込完了";
                    Console.WriteLine(toolStripStatusLabel1.Text);
                    LOG.WriteLine(toolStripStatusLabel1.Text);
                    this.Update();
                    System.Threading.Thread.Sleep(500);
                    return;
                }
                if (endFlag == true)
                {
                    toolStripStatusLabel1.Text = "書込中停止";
                    LOG.WriteLine(toolStripStatusLabel1.Text);
                    DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (result == DialogResult.OK)
                    {
                        serialPort1.Close();
                        LOG.Close();
                        Application.Exit();
                    }
                    return;
                }
            }
            catch
            {
                timeOut = false;
                toolStripStatusLabel1.Text = "書込できません。(その他)";
                Console.WriteLine(toolStripStatusLabel1.Text);
                LOG.WriteLine(toolStripStatusLabel1.Text);
                DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                if (result == DialogResult.OK)
                {
                    serialPort1.Close();
                    LOG.Close();
                    Application.Exit();
                }
            }
        }

        private void SerialWrite(object sender, EventArgs e)
        {

            foreach (string s in RxData)
            {
                if (s.Contains("!!CITYCODE=" + textBox_citycode.Text))
                {
                    Console.WriteLine("!");
                    cityFlag = true;
                    toolStripProgressBar1.Value = 40;
                    return;
                }
                if (s.Contains("!!SENSORNO=" + textBox_devicecode.Text))
                {
                    Console.WriteLine("!");
                    sensFlag = true;
                    toolStripProgressBar1.Value = 50;
                    return;
                }
                if (s.Contains("OK"))
                {
                    if (cityFlag == true)
                    {
                        toolStripProgressBar1.Value = 45;
                        toolStripStatusLabel1.Text = "市町村コード書込";
                        LOG.WriteLine(toolStripStatusLabel1.Text); //CITYCODE
                        cityFlag = false;

                        //次へ進む
                        startUp = true;
                        return;
                    }
                    if (sensFlag == true)
                    {
                        toolStripProgressBar1.Value = 55; //action-10
                        toolStripStatusLabel1.Text = "水位計番号書込";
                        LOG.WriteLine(toolStripStatusLabel1.Text); //SENSORNO
                        sensFlag = false;

                        //次へ進む
                        startUp = true;
                        return;
                    }
                }
            }
        }

        private void CodeTest()
        {
            if (serialPort1.IsOpen)
            {
                serialPort1.DiscardInBuffer();
                serialPort1.DiscardOutBuffer();
            }
            else
            {
                toolStripStatusLabel1.Text = "ケーブルが抜けました。";
                LOG.WriteLine(toolStripStatusLabel1.Text);
                DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                if (result == DialogResult.OK)
                {
                    LOG.Close();
                    Application.Exit();
                }
                return;
            }
            try
            {
                while (startUp == false && endFlag == false && errFlag == false)
                {
                    this.Activate();
                    this.Update();
                    Application.DoEvents();

                    if (timeOut == true)
                    {
                        DateTime endDT = DateTime.Now;
                        TimeSpan ts = endDT - startTime;
                        //Console.WriteLine(ts);
                        if (ts.TotalSeconds > 10)
                        {
                            //タイムアウトした
                            toolStripStatusLabel1.Text = "確認できません。(10秒タイムアウト)";
                            LOG.WriteLine(toolStripStatusLabel1.Text);
                            DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            if (result == DialogResult.OK)
                            {
                                serialPort1.Close();
                                LOG.Close();
                                Application.Exit();
                            }
                            return;
                        }
                    }
                    if (serialPort1.BytesToRead > 0)
                    {
                        dataIN += serialPort1.ReadExisting();
                        if (dataIN.IndexOf("\n") > 0)
                        {
                            RxData = dataIN.Split('\n'); //dataIN.Split(new string[] { "\n" }, StringSplitOptions.None); //Split('\n')
                            foreach (String s in RxData)
                            {
                                Console.WriteLine("LOG: " + s);
                            }
                            this.Invoke(new EventHandler(SerialTest));
                        }
                        dataIN = string.Empty;
                    }
                }
                if (startUp == true)
                {
                    toolStripStatusLabel1.Text = "確認完了";
                    Console.WriteLine(toolStripStatusLabel1.Text);
                    LOG.WriteLine(toolStripStatusLabel1.Text);
                    this.Update();
                    System.Threading.Thread.Sleep(500);
                    return;
                }
                if (endFlag == true)
                {
                    toolStripStatusLabel1.Text = "確認中停止";
                    LOG.WriteLine(toolStripStatusLabel1.Text);
                    DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (result == DialogResult.OK)
                    {
                        serialPort1.Close();
                        LOG.Close();
                        Application.Exit();
                    }
                    return;
                }
                if (errFlag == true)
                {
                    toolStripStatusLabel1.Text = "設定値が水位計IDファイルと異なります。";
                    LOG.WriteLine(toolStripStatusLabel1.Text);
                    DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (result == DialogResult.OK)
                    {
                        serialPort1.Close();
                        LOG.Close();
                        Application.Exit();
                    }
                    return;
                }
            }
            catch
            {
                timeOut = false;
                toolStripStatusLabel1.Text = "確認できません。(その他)";
                Console.WriteLine(toolStripStatusLabel1.Text);
                LOG.WriteLine(toolStripStatusLabel1.Text);
                DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                if (result == DialogResult.OK)
                {
                    serialPort1.Close();
                    LOG.Close();
                    Application.Exit();
                }
            }
        }

        private void SerialTest(object sender, EventArgs e)
        {

            foreach (string s in RxData)
            {
                if (s.Contains("CITYCODE="))
                {
                    Console.WriteLine("!");

                    string str0 = RxData[0];
                    string str1 = RxData[1];


                    if (str0 == "CITYCODE=" + textBox_citycode.Text)
                    {
                        Console.WriteLine("!!");
                        toolStripProgressBar1.Value = 60;
                        toolStripStatusLabel1.Text = "市町村コード確認";
                        LOG.WriteLine(toolStripStatusLabel1.Text);
                    }
                    else
                    {
                        timeOut = false;
                        errFlag = true;
                    }
                    if (str1 == "SENSORNO=" + textBox_devicecode.Text)
                    {
                        Console.WriteLine("!!!");
                        toolStripProgressBar1.Value = 70;
                        toolStripStatusLabel1.Text = "水位計番号確認";
                        LOG.WriteLine(toolStripStatusLabel1.Text);
                        startUp = true;
                        return;
                    }
                    else
                    {
                        timeOut = false;
                        errFlag = true;
                    }
                }
            }
        }

        private void AttachTest()
        {
            if (serialPort1.IsOpen)
            {
                serialPort1.DiscardInBuffer();
                serialPort1.DiscardOutBuffer();
            }
            else
            {
                toolStripStatusLabel1.Text = "ケーブルが抜けました。";
                LOG.WriteLine(toolStripStatusLabel1.Text);
                DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                if (result == DialogResult.OK)
                {
                    LOG.Close();
                    Application.Exit();
                }
                return;
            }
            try
            {
                while (startUp == false && endFlag == false && errFlag == false)
                {
                    this.Activate();
                    this.Update();
                    Application.DoEvents();

                    if (timeOut == true)
                    {
                        DateTime endDT = DateTime.Now;
                        TimeSpan ts = endDT - startTime;
                        //Console.WriteLine(ts);
                        if (ts.TotalSeconds > 45)
                        {
                            //タイムアウトした
                            toolStripStatusLabel1.Text = "基地局確認できません。(45秒タイムアウト)";
                            LOG.WriteLine(toolStripStatusLabel1.Text);
                            DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            if (result == DialogResult.OK)
                            {
                                serialPort1.Close();
                                LOG.Close();
                                Application.Exit();
                            }
                            return;
                        }
                    }
                    if (serialPort1.BytesToRead > 0)
                    {
                        dataIN += serialPort1.ReadExisting();
                        if (dataIN.IndexOf("\n") > 0)
                        {
                            RxData = dataIN.Split('\n'); //dataIN.Split(new string[] { "\n" }, StringSplitOptions.None); //Split('\n')
                            foreach (String s in RxData)
                            {
                                Console.WriteLine("LOG: " + s);
                            }
                            this.Invoke(new EventHandler(SerialAttach));
                        }
                        dataIN = string.Empty;
                    }
                }
                if (startUp == true)
                {
                    toolStripStatusLabel1.Text = "基地局接続確認完了";
                    Console.WriteLine(toolStripStatusLabel1.Text);
                    LOG.WriteLine(toolStripStatusLabel1.Text);
                    this.Update();
                    System.Threading.Thread.Sleep(500);
                    return;
                }
                if (endFlag == true)
                {
                    toolStripStatusLabel1.Text = "基地局接続確認中停止";
                    LOG.WriteLine(toolStripStatusLabel1.Text);
                    DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (result == DialogResult.OK)
                    {
                        serialPort1.Close();
                        LOG.Close();
                        Application.Exit();
                    }
                    return;
                }
                if (errFlag == true)
                {
                    toolStripStatusLabel1.Text = "基地局接続確認できません。(通信部エラー)";
                    LOG.WriteLine(toolStripStatusLabel1.Text);
                    DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (result == DialogResult.OK)
                    {
                        serialPort1.Close();
                        LOG.Close();
                        Application.Exit();
                    }
                    return;
                }
            }
            catch
            {
                timeOut = false;
                toolStripStatusLabel1.Text = "基地局接続確認できません。(その他)";
                Console.WriteLine(toolStripStatusLabel1.Text);
                LOG.WriteLine(toolStripStatusLabel1.Text);　//tama
                DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                if (result == DialogResult.OK)
                {
                    serialPort1.Close();
                    LOG.Close();
                    Application.Exit();
                }
            }
        }

        private void SerialAttach(object sender, EventArgs e)
        {
            foreach (string s in RxData)
            {
                if (s.Contains("!!ATTACH"))
                {
                    Console.WriteLine("!");
                    toolStripProgressBar1.Value = 80;
                    toolStripStatusLabel1.Text = "基地局接続開始 [強制終了45秒]";
                    return;
                }
                if (s.Contains("ATTACH OK"))
                {
                    Console.WriteLine("!");
                    toolStripProgressBar1.Value = 90;
                    startUp = true;
                    return;
                }
                else if(s.Contains("ATTACH ERROR"))
                {
                    timeOut = false;
                    errFlag = true;
                }

            }
        }
        private void WriteCommand(string CMD)
        {
            switch(CMD)
            {
                case "CITYCODE":
                    string CCODE = "!!CITYCODE=" + textBox_citycode.Text + "\r\n";
                    serialPort1.WriteLine(CCODE);
                    timeOut = true;
                    startTime = DateTime.Now;
                    break;

                case "SENSORNO":
                    string SCODE = "!!SENSORNO=" + textBox_devicecode.Text + "\r\n";
                    serialPort1.WriteLine(SCODE);
                    timeOut = true;
                    startTime = DateTime.Now;
                    break;

                case "ATTACH":
                    string ATT = "!!ATTACH" + Environment.NewLine;
                    serialPort1.WriteLine(ATT);
                    timeOut = true;
                    startTime = DateTime.Now;
                    break;

                case "INFO":
                    string INF = "!!INFO" + Environment.NewLine;
                    serialPort1.WriteLine(INF);
                    timeOut = true;
                    startTime = DateTime.Now;
                    break;
            }
        }

        public static string[,] GetDeviceNames()
        {

            var deviceNameList = new System.Collections.ArrayList();
            var deviceIDList = new System.Collections.ArrayList();
            var check = new System.Text.RegularExpressions.Regex("(COM[1-9][0-9]?[0-9]?)");

            ManagementClass mclass = new ManagementClass("Win32_PnPEntity");
            ManagementObjectCollection manageObjCol = mclass.GetInstances();

            //全てのPnPデバイスを探索しシリアル通信が行われるデバイスを随時追加する
            foreach (ManagementObject manageObj in manageObjCol)
            {
                //Nameプロパティを取得
                var namePropertyValue = manageObj.GetPropertyValue("Name");
                if (namePropertyValue == null)
                {
                    continue;
                }

                //DeviceIDプロパティを取得
                var deviceIDPropertyValue = manageObj.GetPropertyValue("DeviceID");
                if (deviceIDPropertyValue == null)
                {
                    continue;
                }

                //Nameプロパティ文字列の一部が"(COM1)～(COM999)"と一致するときリストに追加"
                string name = namePropertyValue.ToString();
                string deviceID = deviceIDPropertyValue.ToString();
                if (check.IsMatch(name))
                {
                    deviceNameList.Add(name);
                    deviceIDList.Add(deviceID);
                }
            }

            //戻り値作成
            if (deviceNameList.Count > 0)
            {
                string[,] deviceNames = new string[deviceNameList.Count, 2];
                var name = deviceNameList;
                var ID = deviceIDList;



                for(int i=0; i < deviceNameList.Count; i++)
                {
                    deviceNames[i, 0] = name[i].ToString();
                    deviceNames[i, 1] = ID[i].ToString();
                }

                /*
                int index = 0;
                foreach (var name in deviceNameList)
                {
                    string dev = name.ToString();
                    //dev = dev.Substring(0, dev.IndexOf("("));
                    deviceNames[index++,] = dev;
                }*/

                return deviceNames;
            }
            else
            {
                return null;
            }
        }

        private void バージョンToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("危機管理型水位計MW-001 水位計ID書換ツール\nバージョン: 1.00");
        }
    }
}
