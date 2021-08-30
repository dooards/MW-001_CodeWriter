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
        //string[] RxData;
        string dataIN;
        string OldText;
        string[] GenCable= new string[10];
        string[] FDTI = {"FT4TWAWE", "FT4XOY0Q", "FT4TYU5R", "FT4TZ2NT", "FT4XSQD3", "FT4U1Z7P", "FT4XOWVF", "FT4XSQ4O", "FT4TZ322", "FT4XRUHU", "FT4U1V7R", "FT4U1Y6Y" };
        
        


        bool errFlag = false; //error発生フラグ
        bool endFlag = false; //停止ボタンフラグ
        bool statusFlag = false; //水位計ID設定可
        bool timeOut = false;　//タイムアウトフラグ
        bool startUp = false; //起動済み確認フラグ
        bool idleFlag = false;　//動作待ちフラグ
        bool regularMode = false; //通常モード
        bool cityFlag = false;
        bool sensFlag = false;

        public const int WM_DEVICECHANGE = 0x00000219;  //デバイス変化のWindowsイベントの値



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
            toolStripProgressBar1.Value = 0;
            button_action.Text = "選択";
            button_action.Enabled = false;
            button_stop.Text = "停止";
            textBox_csv.Clear();
            textBox_devicecode.Clear();
            textBox_citycode.Clear();
            textBox_tellnumber.Clear();
            timeLabel.Text = "";
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
            MainThread();
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            switch (m.Msg)
            {
                case WM_DEVICECHANGE:   //デバイス状況の変化イベント
                    // AddMessage(m.ToString() + Environment.NewLine);
                    Task.Run(() => PortIdentify());      //デバイスをチェック
                    break;
            }
        }

        private void CSVSearch()
        {
            //File Path
            string appPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            appPath = appPath.Replace("水位計ID設定ツール.exe", "");
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


            if (errFlag == true)
            {

                DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                if (result == DialogResult.OK)
                {
                    LOG.Close();
                    Application.Exit();
                }
                return;
            }
            else
            {
                idleFlag = true;
            }

        }

        private void PortIdentify()
        {
            if (serialPort1.IsOpen)
            {
                return;
            }
            else if(idleFlag == false)
            {
                return;
            }
            else
            {
                if (statusFlag == true)
                {
                    toolStripStatusLabel1.Text = "ケーブルが抜けました。";
                    LOG.WriteLine(toolStripStatusLabel1.Text);
                    DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    if (result == DialogResult.OK)
                    {
                        LOG.Close();
                        Application.Exit();
                        endFlag = true;
                    }
                }

                else
                {
                    Invoke(new Action(MainThread));
                }
            }
        }

        public void MainThread()
        {
            if (errFlag == true)
            {
                return;
            }
            else
            {
                idleFlag = true;
            }

            comboBox_comport.Items.Clear();
            string[] ports;
            int Gc = 0;
            ports = SerialPort.GetPortNames();
            if (ports.Length > 0) //COMポートの認識
            {
                portNames = GetDeviceNames();
                for (int i = 0; i < ports.Length; i++)
                {
                    for(int CableNum = 0; CableNum < FDTI.Length; CableNum++)
                    {
                        string Serial = FDTI[CableNum];
                        if (portNames[i, 1].Contains(Serial)) //純正ケーブル
                        {
                            GenCable[Gc] = portNames[i, 0];
                            Gc++;

                            /*comboBox_comport.Items.Add("MW-001用接続ケーブル");
                            int num = comboBox_comport.Items.Count - 1;
                            comboBox_comport.SelectedIndex = num;
                            
                           
                            toolStripStatusLabel1.Text = "ケーブル選択可能";
                            toolStripProgressBar1.Value = 10;
                            button_action.Enabled = true;
                            */
                            
                        }
                        else //純正外
                        {
                            /*
                            //comboBox_comport.SelectedIndex = -1;
                            //comboBox_comport.Text = "";
                            toolStripStatusLabel1.Text = "ケーブル検索中";
                            toolStripProgressBar1.Value = 5;
                            button_action.Enabled = false;
                            */
                        }
                        

                    }
                }
                if(Gc == 0)
                {
                    toolStripStatusLabel1.Text = "ケーブル検索中";
                    toolStripProgressBar1.Value = 5;
                    button_action.Enabled = false;
                }
                else
                {
                    for(int num = 0; num < Gc; num++)
                    {
                        comboBox_comport.Items.Add("MW-001用接続ケーブル "); //  + num.ToString());
                    }
                    comboBox_comport.SelectedIndex = Gc -1;
                    toolStripStatusLabel1.Text = "ケーブル選択可能";
                    toolStripProgressBar1.Value = 10;
                    button_action.Enabled = true;
                }

            }
            else //未接続
            {
                //comboBox_comport.SelectedIndex = -1;
                comboBox_comport.Text = "";

                toolStripStatusLabel1.Text = "ケーブル検索中";
                toolStripProgressBar1.Value = 5;
                button_action.Enabled = false;
            }

        }

        private void button_stop_Click(object sender, EventArgs e)
        {
            endFlag = true;
            
            OldText = toolStripStatusLabel1.Text;
            toolStripStatusLabel1.Text = "停止ボタン";
            LOG.WriteLine(toolStripStatusLabel1.Text);

            if (idleFlag == true)
            {
                if(startUp == true)
                {
                    toolStripStatusLabel1.Text = "終了します。";
                    //DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "終了", MessageBoxButtons.OK, MessageBoxIcon.None);

                    //if (result == DialogResult.OK)
                    //{
                    LOG.WriteLine("完了");
                        LOG.Close();
                        Application.Exit();
                    //}
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
                        toolStripStatusLabel1.Text = OldText;
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
                
                button_action.Text = "設定";
                NextPannel();

                toolStripStatusLabel1.Text = "所定の位置に磁石を近づけて、電源を入れて下さい。";
                regularMode = true;
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
                    toolStripStatusLabel1.Text = "水位計IDの設定が可能です。";
                }
                return;

            }
            if(statusFlag == true)
            {
                button_action.Enabled = false;
                idleFlag = false;
                
                WriteCommand("AT");
                ATTest();

                if (startUp == true)
                {
                    startUp = false;
                    WriteCommand("CITYCODE");
                    CodeWrite();
                }

                if (startUp == true)
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

                System.Threading.Thread.Sleep(250);

                if (startUp == true)
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
                    toolStripStatusLabel1.Text = "水位計ID設定完了 終了ボタンを押して終了してください。";
                    LOG.WriteLine(toolStripStatusLabel1.Text);
                    return;
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
                string strValue = GenCable[comboBox_comport.SelectedIndex];
                strValue = strValue.Remove(0, strValue.IndexOf("(") + 1);
                strValue = strValue.Remove(strValue.IndexOf(")"));

                serialPort1.PortName = strValue;
                serialPort1.BaudRate = 115200; // Convert.ToInt32(cBoxBAUDRATE.Text);
                serialPort1.DataBits = 8; // Convert.ToInt32(cBoxDATABITS.Text);
                serialPort1.StopBits = (StopBits)Enum.Parse(typeof(StopBits), "One");
                serialPort1.Parity = (Parity)Enum.Parse(typeof(Parity), "None");

                serialPort1.DiscardNull = true;
                serialPort1.NewLine = "\n";
                //serialPort1.DataReceived += new SerialDataReceivedEventHandler(serialPort1_DataReceived);

                serialPort1.Open();
                if (serialPort1.IsOpen)
                {
                    toolStripStatusLabel1.Text = "接続済み";
                    LOG.WriteLine("SELECT:" + GenCable[comboBox_comport.SelectedIndex]); //COM名
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
                toolStripStatusLabel1.Text = "ケーブルが抜けました。";
                LOG.WriteLine(toolStripStatusLabel1.Text);
                DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                if (result == DialogResult.OK)
                {
                    LOG.Close();
                    Application.Exit();
                }
            }
        }

        private void NextPannel()
        {
            panel1.Visible = false;
            panel2.Visible = true;
        }

        /*
        private void serialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            string str = serialPort1.ReadLine();
            Console.WriteLine(str);
        }
        */

        private void PowerON()
        {
            try
            {
                while (startUp == false  && errFlag == false && endFlag == false)
                {
                    this.Activate();
                    //this.Update();
                    Application.DoEvents();
                    
                    if (timeOut == true)
                    {
                        DateTime endDT = DateTime.Now;
                        TimeSpan ts = endDT - startTime;
                        timeLabel.Text = Math.Floor(ts.TotalSeconds).ToString();
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
                        /*dataIN += serialPort1.ReadLine();
                        if(dataIN.IndexOf("\n") > 0)
                        {
                            RxData = dataIN.Split('\n'); //dataIN.Split(new string[] { "\n" }, StringSplitOptions.None); //Split('\n')
                            foreach (string s in RxData)
                            {
                                Console.WriteLine("LOG: " + s);
                            }
                            this.Invoke(new EventHandler(SerialLog));
                        }
                        dataIN = string.Empty;*/

                        dataIN = serialPort1.ReadLine();
                        Console.WriteLine("LOG: " + dataIN);
                        this.Invoke(new EventHandler(SerialLog));

                    }

                    if (endFlag == true)
                    {
                        if (timeOut == true)
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
                            break;
                        }
                        else
                        {
                            //OldText = toolStripStatusLabel1.Text;
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
                                toolStripStatusLabel1.Text = OldText;
                                LOG.WriteLine("No");
                                endFlag = false;
                            }
                        }
                    }

                }
                if (startUp == true)
                {
                    toolStripStatusLabel1.Text = "起動完了";
                    Console.WriteLine(toolStripStatusLabel1.Text);
                    LOG.WriteLine(toolStripStatusLabel1.Text);
                    timeLabel.Text = "";
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
                toolStripStatusLabel1.Text = "プログラム動作中エラー";
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
            //foreach (string s in RxData)
            string s = dataIN;
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
                    regularMode = false;
                    toolStripProgressBar1.Value = 20;
                    startTime = DateTime.Now; //時間取得
                    Console.WriteLine("LOG: " + startTime);
                    timeOut = true; //タイマー起動
                    toolStripStatusLabel1.Text = "制御部 起動開始 [強制終了40秒]";
                    Console.WriteLine(toolStripStatusLabel1.Text);
                    LOG.WriteLine(toolStripStatusLabel1.Text); 
                    return;
                }

                if (s.Contains("WAKEUP"))
                {
                    if(regularMode == false)
                    {
                        toolStripProgressBar1.Value = 25;
                        toolStripStatusLabel1.Text = "通信部 起動開始 [強制終了40秒]";
                        Console.WriteLine(toolStripStatusLabel1.Text);
                        LOG.WriteLine(toolStripStatusLabel1.Text); //テストモード起動中
                        return;
                    }
                }

                if (s.Contains("AT%"))
                {
                    if(regularMode == true)
                    {
                        toolStripStatusLabel1.Text = "テストモードで起動していません。水位計の電源を落としてください。";
                        errFlag = true;
                    }
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

                        if (!callnum.StartsWith("0"))
                        {
                            callnum = "0" + callnum;
                        }


                        if (textBox_tellnumber.Text == callnum)
                        {
                            if(textBox_csv.Text.Substring(0, textBox_csv.Text.Length-4) == sbuf[1])
                            {
                                textBox_citycode.Text = sbuf[1];
                                textBox_devicecode.Text = sbuf[2];
                                textBox_name.Text = sbuf[3];
                                toolStripProgressBar1.Value = 40;

                                //水位計ID設定可
                                statusFlag = true;
                                break;
                            }
                            else
                            {
                                statusFlag = false;
                            }

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
                        LOG.WriteLine("自治体コード: " + textBox_citycode.Text);
                        LOG.WriteLine("水位計番号: " + textBox_devicecode.Text);
                        LOG.WriteLine("水位計名称: " + textBox_name.Text);
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



        private void ATTest()
        {
            dataIN = string.Empty;
            try
            {
                while (startUp == false && endFlag == false)
                {
                    this.Activate();
                    //this.Update();
                    Application.DoEvents();

                    if (timeOut == true)
                    {
                        DateTime endDT = DateTime.Now;
                        TimeSpan ts = endDT - startTime;
                        //Console.WriteLine(ts);
                        if (ts.TotalSeconds > 10)
                        {
                            //タイムアウトした
                            toolStripStatusLabel1.Text = "機器を認識できません。(10秒タイムアウト)";
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
                        dataIN += serialPort1.ReadLine();/*
                        if (dataIN.IndexOf("\n") > 0)
                        {
                            RxData = dataIN.Split('\n'); //dataIN.Split(new string[] { "\n" }, StringSplitOptions.None); //Split('\n')
                            foreach (string s in RxData)
                            {
                                Console.WriteLine("LOG: " + s);
                            }
                            this.Invoke(new EventHandler(SerialAT));
                        }
                        dataIN = string.Empty;*/
                        //dataIN = serialPort1.ReadLine();
                        Console.WriteLine("LOG: " + dataIN);
                        this.Invoke(new EventHandler(SerialAT));
                    }
                }
                if (startUp == true)
                {
                    toolStripStatusLabel1.Text = "機器認識済み";
                    Console.WriteLine(toolStripStatusLabel1.Text);
                    LOG.WriteLine(toolStripStatusLabel1.Text);
                    this.Update();
                    System.Threading.Thread.Sleep(250);
                    return;
                }
                if (endFlag == true)
                {
                    toolStripStatusLabel1.Text = "機器認識できず";
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
                toolStripStatusLabel1.Text = "プログラム動作中エラー（認識中）";
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

        private void SerialAT(object sender, EventArgs e)
        {

            //foreach (string s in RxData)
            string s = dataIN;
            {

                if (s.Contains("OK"))
                {
                    //次へ進む
                    startUp = true;
                    return;
                }
            }
        }

        private void CodeWrite()
        {
            try
            {
                while (startUp == false && endFlag == false)
                {
                    this.Activate();
                    //this.Update();
                    Application.DoEvents();

                    if (timeOut == true)
                    {
                        DateTime endDT = DateTime.Now;
                        TimeSpan ts = endDT - startTime;
                        //Console.WriteLine(ts);
                        if (ts.TotalSeconds > 10)
                        {
                            //タイムアウトした
                            toolStripStatusLabel1.Text = "設定できません。(10秒タイムアウト)";
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
                        /*
                        dataIN += serialPort1.ReadExisting();//ReadExisting();

                        if(dataIN.IndexOf("\n") > 0)
                        {
                            RxData = dataIN.Split('\n'); //dataIN.Replace("\r\n", "").Split('\n');
                            foreach (string s in RxData)
                            {
                                Console.WriteLine("LOG: " + s);
                            }
                            this.Invoke(new EventHandler(SerialWrite));
                        }
                        dataIN = string.Empty;
                        */
                        dataIN = serialPort1.ReadLine();
                        Console.WriteLine("LOG: " + dataIN);
                        this.Invoke(new EventHandler(SerialWrite));
                    }

                }
                if (startUp == true)
                {
                    toolStripStatusLabel1.Text = "設定完了";
                    Console.WriteLine(toolStripStatusLabel1.Text);
                    LOG.WriteLine(toolStripStatusLabel1.Text);
                    this.Update();
                    System.Threading.Thread.Sleep(250);
                    return;
                }
                if (endFlag == true)
                {
                    toolStripStatusLabel1.Text = "設定中停止";
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
                toolStripStatusLabel1.Text = "プログラム動作中エラー（設定中）";
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

            //foreach (string s in RxData)
            string s = dataIN;
            {
                if (s.Contains("!!CITYCODE=" + textBox_citycode.Text))
                {
                    cityFlag = true;
                    toolStripProgressBar1.Value = 40;
                    return;
                }
                if (s.Contains("!!SENSORNO=" + textBox_devicecode.Text))
                {
                    sensFlag = true;
                    toolStripProgressBar1.Value = 50;
                    return;
                }
                if (s.Contains("OK"))
                {
                    if (cityFlag == true)
                    {
                        toolStripProgressBar1.Value = 45;
                        toolStripStatusLabel1.Text = "自治体コード設定";
                        LOG.WriteLine(toolStripStatusLabel1.Text); //CITYCODE
                        cityFlag = false;

                        //次へ進む
                        startUp = true;
                        return;
                    }
                    if (sensFlag == true)
                    {
                        toolStripProgressBar1.Value = 55; //action-10
                        toolStripStatusLabel1.Text = "水位計番号設定";
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
            try
            {
                while (startUp == false && endFlag == false && errFlag == false)
                {
                    this.Activate();
                    //this.Update();
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
                    if (serialPort1.BytesToRead > 2)
                    {
                        /*
                        dataIN += serialPort1.ReadExisting();
                        if (dataIN.IndexOf("\n") > 0)
                        {
                            RxData = dataIN.Split('\n'); //dataIN.Split(new string[] { "\n" }, StringSplitOptions.None); //Split('\n')
                            foreach (string s in RxData)
                            {
                                Console.WriteLine("LOG: " + s);
                            }
                            this.Invoke(new EventHandler(SerialTest));
                        }
                        dataIN = string.Empty;
                        */
                        dataIN = serialPort1.ReadLine();
                        Console.WriteLine("LOG: " + dataIN);
                        this.Invoke(new EventHandler(SerialTest));
                    }

                }
                if (startUp == true)
                {
                    toolStripStatusLabel1.Text = "確認完了";
                    Console.WriteLine(toolStripStatusLabel1.Text);
                    LOG.WriteLine(toolStripStatusLabel1.Text);
                    this.Update();
                    System.Threading.Thread.Sleep(250);
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
                toolStripStatusLabel1.Text = "プログラム動作中エラー（確認中）";
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

            //foreach (string s in RxData)
            string s = dataIN;
            {
                if (s.Contains("CITYCODE="))
                {
                    //string str0 = RxData[0];
                    //string str1 = RxData[1];

                    //if (str0 == "CITYCODE=" + textBox_citycode.Text)
                    if (s == "CITYCODE=" + textBox_citycode.Text)
                    {
                        toolStripProgressBar1.Value = 60;
                        toolStripStatusLabel1.Text = "自治体コード確認";
                        LOG.WriteLine(toolStripStatusLabel1.Text);
                    }
                    else
                    {
                        timeOut = false;
                        errFlag = true;
                    }
                }
                if (s.Contains("SENSORNO="))
                {
                    //if (str1 == "SENSORNO=" + textBox_devicecode.Text)
                    if (s == "SENSORNO=" + textBox_devicecode.Text)
                    {
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
                        timeLabel.Text = Math.Floor(ts.TotalSeconds).ToString();
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
                    if (serialPort1.BytesToRead > 2)
                    {
                        /*dataIN += serialPort1.ReadExisting();
                        if (dataIN.IndexOf("\n") > 0)
                        {
                            RxData = dataIN.Split('\n'); 
                            foreach (string s in RxData)
                            {
                                Console.WriteLine("LOG: " + s);
                            }
                            this.Invoke(new EventHandler(SerialAttach));
                        }
                        dataIN = string.Empty;
                        */
                        dataIN = serialPort1.ReadLine();
                        Console.WriteLine("LOG: " + dataIN);
                        this.Invoke(new EventHandler(SerialAttach));
                    }

                }
                

                if (startUp == true)
                {
                    toolStripStatusLabel1.Text = "基地局接続確認完了";
                    Console.WriteLine(toolStripStatusLabel1.Text);
                    LOG.WriteLine(toolStripStatusLabel1.Text);
                    timeLabel.Text = "";
                    this.Update();
                    System.Threading.Thread.Sleep(250);
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
                    //toolStripStatusLabel1.Text = "基地局接続確認できません。(通信部エラー)";
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
                toolStripStatusLabel1.Text = "プログラム動作中エラー（基地局接続確認中）";
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
            //foreach (string s in RxData)
            string s = dataIN;
            {
                if (s.Contains("!!ATTACH"))
                {
                    toolStripProgressBar1.Value = 80;
                    toolStripStatusLabel1.Text = "基地局接続開始 [強制終了45秒]";
                    return;
                }
                if (s.Contains("OK"))
                {
                    toolStripProgressBar1.Value = 90;
                    startUp = true;
                    return;
                }
                else if(s.Contains("ATTACH ERROR"))
                {
                    timeOut = false;
                    errFlag = true;
                }
                if (s.Contains("***"))
                {
                    toolStripStatusLabel1.Text = "基地局接続確認できません。(リセット)";
                    errFlag = true;
                    return;
                }

            }
        }
        private void WriteCommand(string CMD)
        {
            try
            {
                serialPort1.DiscardInBuffer();
                serialPort1.DiscardOutBuffer();

                switch (CMD)
                {
                    case "AT":
                        string ATCom = "!!AT" + Environment.NewLine;
                        serialPort1.WriteLine(ATCom);
                        timeOut = true;
                        startTime = DateTime.Now;
                        break;

                    case "CITYCODE":
                        string CCODE = "!!CITYCODE=" + textBox_citycode.Text + Environment.NewLine;
                        serialPort1.WriteLine(CCODE);
                        timeOut = true;
                        startTime = DateTime.Now;
                        break;

                    case "SENSORNO":
                        string SCODE = "!!SENSORNO=" + textBox_devicecode.Text + Environment.NewLine;
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
            catch
            {
                toolStripStatusLabel1.Text = "コマンド送信できません。";
                Console.WriteLine(toolStripStatusLabel1.Text);
                LOG.WriteLine(toolStripStatusLabel1.Text);
                DialogResult result = MessageBox.Show(toolStripStatusLabel1.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                if (result == DialogResult.OK)
                {

                }
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

                return deviceNames;
            }
            else
            {
                return null;
            }
        }

        private void バージョンToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("危機管理型水位計MW-001 水位計ID設定ソフトウェア\nバージョン: 1.00\n" +
                "Copyright (c) 2021 ABIT Co.\nReleased under the MIT license\n" +
                "https://opensource.org/licenses/mit-license.php");
        }
    }
}
