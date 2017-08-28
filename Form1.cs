using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using clsDBMySql;
using System.Threading;
using clsPLC;
using Utilities;
using clsDBMySql3;
using nsPlcHelper;
using ATLocv.MiDataCollectForProcessLotForEach;

namespace ATLocv
{
    public partial class Form1 : Form
    {
        bool webtrue = true;//是否连接WEB数据库成功
        int iLive = 0;
        int isScanned = 0;
        int isWriteType = 0;
        int isUpload = 0;
        int NGCount = 0;
        string trayCode = "";
        int rowsCount = 0;
        int count = 0;
        int count2 = 0;
        int needleCount = 0;
        int isClear = 0;
        int type = -1;
        int adjCount = 0;
        float masterSum = 0;
        float master;
        float v;
        float r;
        float temp;
        double timeRes;
        int isFangdai = 0;
        //int iPutFloorStatus = 0;
        //int iOutFloorStatus = 0;
        static string strCurMsg = "";
        string strIP = "192.168.30.152";
        int iPort = 5025;
        string strIP2 = "192.168.1.155";//PLC的IP
        int iPort2 = 9094;//PLC的端口
        string strStrartIime = "";
        string nextAdjDateTime = "";
        int iAdjOffHour = 8;
        int PutStep = 0;
        //string intime = "";
        //string outtime = "";
        static object obj = new object();
        TextBox[] o1 = null;//存放O1测试结果
        TextBox[] ob = null;//存放OB测试结果
        TextBox[] sun = null;//存放O1OB测试总结果
        Label[] lb = null;//存放label控件
        TextBox[] tb = null;//存放textBox控件（模拟托盘）
        List<string> lstStrMsg = new List<string>();//存放消息提示的泛型
        clsDBMysql mSql = new clsDBMysql();//数据库150连接类
        clsDBMysql3 mysql3 = new clsDBMysql3();//WEB端数据库连接类
        clientTest ctTest1 = null;//读取电压表数据的类
        cPlcTcpipComm plcComm = new cPlcTcpipComm();
        clientTest plcNetOperator = null;//与PLC通信的类
        clientTest scannerConn = new clientTest("192.168.30.125", 1124);
        //clsPLC_sx plcOperator = null;//plc读写数据的相关类
        Thread thOCVLive = null;//定义OCV线程
        IniFile iniFile = new IniFile(Application.StartupPath + "\\config.ini");
        float[] dianya = new float[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
        float[] naizu = new float[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
        float[] wendu = new float[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
        float[] shellVol = new float[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
        string[] jg = new string[] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
        Dictionary<string, string> pList = new Dictionary<string, string>();
        Dictionary<int, int> dicReq = new Dictionary<int, int>();//存放请求
        Dictionary<int, int> dicFloorStatus = new Dictionary<int, int>();//存放层状态

        public Form1()
        {
            InitializeComponent();
            lb = new Label[] { l1, l2, l3, l4, l5, l6, l7, l8, l9, l10, l11, l12, l13, l14, l15, l16, l17, l18, l19, l20, l21, l22, l23, l24 };
            tb = new TextBox[] { tb1, tb2, tb3, tb4, tb5, tb6, tb7, tb8, tb9, tb10, tb11, tb12, tb13, tb14, tb15, tb16, tb17, tb18, tb19, tb20, tb21, tb22, tb23, tb24 };
            o1 = new TextBox[] { tbo11, tbo12, tbo13, tbo14, tbo15, tbo16 };
            ob = new TextBox[] { tbob1, tbob2, tbob3, tbob4, tbob5, tbob6 };
            sun = new TextBox[] { tbs1, tbs2, tbs3, tbs4, tbs5, tbs6 };
            canshuJiazai();
            for (int i = 0; i < 24; i++)
            {
                lb[i].Text = (i + 1).ToString("00");
            }
            timer1.Enabled = true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dd();
            UpdateOCVMaster(float.Parse(tbMaster.Text), float.Parse(tbMasterOffer.Text));
        }

        private void Form1Closed(object sender, FormClosedEventArgs e)
        {
            plcComm.SetDevice("R0332", 0);
            plcComm.SetDevice("R0270", 0);
        }

        public void dd()
        {
            // 写入ini
            //Ini ini = new Ini(Application.StartupPath + "\\Mesconfig.ini");
            //ini.Writue("Setting", "key1", " WORLD!");
            //ini.Writue("Setting", "key2", "HELLO CHINA!");

            // 读取ini
            Ini ini = new Ini(Application.StartupPath + "\\Mesconfig.ini");
            string str1 = ini.ReadValue("miDataCollectForProcessLotForEach1", "DataCollectForProcessLotForEachUrl");
            string str2 = ini.ReadValue("miDataCollectForProcessLotForEach1", "DataCollectForProcessLotForEachSite");
            string str3 = ini.ReadValue("miDataCollectForProcessLotForEach1", "DataCollectForProcessLotForEachOperation");
            string str4 = ini.ReadValue("miDataCollectForProcessLotForEach1", "DataCollectForProcessLotForEachOperationRevision");
            string str5 = ini.ReadValue("miDataCollectForProcessLotForEach1", "DataCollectForProcessLotForEachResource");
            string str6 = ini.ReadValue("miDataCollectForProcessLotForEach1", "DataCollectForProcessLotForEachUser");
            string str7 = ini.ReadValue("miDataCollectForProcessLotForEach1", "DataCollectForProcessLotForEachActivityID");
            string str8 = ini.ReadValue("miDataCollectForProcessLotForEach1", "DataCollectForProcessLotForEachIsDispositionRequired");
            string str9 = ini.ReadValue("getCellTestResultByTrayId", "CellTestResultByTrayUrl");
            string str10 = ini.ReadValue("getCellTestResultByTrayId", "CellTestResultByTraySite");
            string str11 = ini.ReadValue("getCellTestResultByTrayId", "CellTestResultByTrayOperation");
            string str12 = ini.ReadValue("getCellTestResultByTrayId", "CellTestResultByTrayOperationRevision");
            string str13 = ini.ReadValue("getCellTestResultByTrayId", "CellTestResultByTrayResource");
        }


        #region  数据加载 逻辑不需更改，部分参数可能需要更改
        private void canshuJiazai()
        {
            strIP = iniFile.GetString("XITONG", "34461IP", "");
            iPort = int.Parse(iniFile.GetString("XITONG", "34461Port", "").ToString());
            strIP2 = iniFile.GetString("XITONG", "plcIP", "");
            iPort2 = int.Parse(iniFile.GetString("XITONG", "plcPort", "").ToString());
            textBox6.Text = iniFile.GetString("XITONG", "34461IP", "");
            textBox7.Text = iniFile.GetString("XITONG", "34461Port", "");
            nextAdjDateTime = iniFile.GetString("CLEAR", "time", "");
            strStrartIime = iniFile.GetString("CLEAR", "starttime", "09:00");
            iAdjOffHour = int.Parse(iniFile.GetString("CLEAR", "offtime", "8"));
            sp3562.PortName = (iniFile.GetString("XITONG", "sp3562", "")).ToString();
            sptemp.PortName = iniFile.GetString("XITONG", "cewen", "");
            sptemp2.PortName = iniFile.GetString("XITONG", "cewen2", "");
            //spTrayIn.PortName = (iniFile.GetString("sys", "spTray", "")).ToString();
            //spTrayOut.PortName = (iniFile.GetString("sys", "spTrayOut", "")).ToString();
            //cgVar.mLoad.sScan1Type = (iniFile.GetString("SYSTEM", "scan1type", "")).ToString();
            //cgVar.mLoad.sScan3Type = (iniFile.GetString("sys", "scan3type", "")).ToString();

            chkMaster.Checked = int.Parse(iniFile.GetString("XITONG", "master", "")) == 1 ? true : false;
            tbvd1.Text = iniFile.GetString("XITONG", "vbu", "");
            tbrd1.Text = iniFile.GetString("XITONG", "masterChazhi", "");
            tbMaster.Text = iniFile.GetString("XITONG", "Mzhi", "");
            tbMasterOffer.Text = iniFile.GetString("XITONG", "Mwucha", "");
            label24.Text = "探针下压次数：" + iniFile.GetString("XITONG", "needleTimes", "");
            needleCount = int.Parse(iniFile.GetString("XITONG", "needleTimes", ""));
            master = float.Parse(iniFile.GetString("XITONG", "masterChazhi", ""));
            txtO1A.Text = iniFile.GetString("TIAOJIAN", "MINO1VOL", "");
            txtO1B.Text = iniFile.GetString("TIAOJIAN", "MAXO1VOL", "");
            txtOBA.Text = iniFile.GetString("TIAOJIAN", "MINOBVOL", "");
            txtOBB.Text = iniFile.GetString("TIAOJIAN", "MINOBVOL", "");
            txtR1.Text = iniFile.GetString("TIAOJIAN", "MINR", "");
            txtR2.Text = iniFile.GetString("TIAOJIAN", "MAXR", "");
            txtShell1.Text = iniFile.GetString("TIAOJIAN", "MINSHELL", "");
            txtShell2.Text = iniFile.GetString("TIAOJIAN", "MAXSHELL", "");
            txbT1.Text = iniFile.GetString("TIAOJIAN", "MINT", "");
            txbT2.Text = iniFile.GetString("TIAOJIAN", "MAXT", "");
            txbNG.Text = iniFile.GetString("TIAOJIAN", "NGNUM", "");
            tbTray.Text = iniFile.GetString("shuju", "tray", "");
            textBox5.Text = iniFile.GetString("shuju", "type", "");
            PutStep = int.Parse(iniFile.GetString("sys", "putstep", ""));
            for (int i = 0; i < 24; i++)
            {
                dianya[i] = float.Parse(iniFile.GetString("CESHI", "v" + (i + 1), ""));
                naizu[i] = float.Parse(iniFile.GetString("CESHI", "r" + (i + 1), ""));
                wendu[i] = float.Parse(iniFile.GetString("CESHI", "t" + (i + 1), ""));
                jg[i] = iniFile.GetString("CESHI", "jieguo" + (i + 1), "");
                string dianchi = "batty" + (i + 1) + "";
                string icxs = "bx" + (i + 1);
                string icxs2 = "lx" + (i + 1);
                tb[i].Text = iniFile.GetString("SHUJU", dianchi, "");
                tb[i].BackColor = iniFile.GetString("SHUJU", icxs, "") == "0" ? System.Drawing.SystemColors.Control : iniFile.GetString("SHUJU", icxs, "") == "1" ? Color.Green : Color.Red;
                lb[i].BackColor = iniFile.GetString("SHUJU", icxs2, "") == "0" ? Color.Green : iniFile.GetString("SHUJU", icxs2, "") == "1" ? Color.Red : System.Drawing.SystemColors.Control;
            }
            for (int i = 0; i < 6; i++)
            {
                o1[i].Text = iniFile.GetString("TONGJI", "O1" + i, "");
                ob[i].Text = iniFile.GetString("TONGJI", "OB" + i, "");
            }
            textBox1.Text = iniFile.GetString("TONGJI", "shijian", "");
        }
        #endregion



        #region 消息提示/文件记录/OCV启动线程/写入数据/导出表格/生产统计数据/统计写入配置文件
        private void ShowErrMsg(string strMsg)
        {

            string str = "[" + DateTime.Now.ToString("HH:mm:ss") + "]:";
            str += strMsg;
            strCurMsg = strMsg;
            if (lstStrMsg.Count == 0)
            {
                lstStrMsg.Insert(0, str);
                SaveOCVLog(str);
            }
            else
            {
                if (lstStrMsg.Count > 20)
                    lstStrMsg.RemoveAt(lstStrMsg.Count - 1);
                int iCount = lstStrMsg.Count;
                if (iCount > 2)
                    iCount = 2;
                int iRtn = lstStrMsg.FindIndex(0, iCount, FindMsgIndex);
                if (iRtn == -1)
                {
                    if (webtrue)
                    {
                        System.Guid guid = new Guid();
                        guid = Guid.NewGuid();
                        //随机生成如下字符串(32位,保证ID不重复)：
                        //e92b8e30-a6e5-41f6-a6b9-188230a23dd2
                        AddFacility_run_information(guid.ToString(), 8, 1, strMsg, DateTime.Now);
                    }
                    lstStrMsg.Insert(0, str);
                    SaveOCVLog(str);
                    lstMsg.Items.Clear();
                    lstMsg.Items.AddRange(lstStrMsg.ToArray());
                }
                else
                {
                    lstStrMsg.RemoveAt(iRtn);
                    lstStrMsg.Insert(0, str);
                    lstMsg.Items.Clear();
                    lstMsg.Items.AddRange(lstStrMsg.ToArray());
                }
                
            }
            //if (lstMsg.Items.Count > 30)
            //{
            //    int indexCount = lstMsg.Items.Count - 30;
            //    if (indexCount == 1)
            //        lstMsg.Items.Remove(0);
            //    else
            //    {
            //        for (int i = 0; i < indexCount; i++)
            //        {
            //            lstMsg.Items.Remove(indexCount);
            //        }
            //    }
            //    //lstMsg.Items.Clear();
            //    //lstMsg.Items.AddRange(lstStrMsg.ToArray());
            //}
        }
        private void SaveOCVLog(string msg)
        {
            if (!Directory.Exists(Application.StartupPath + "\\" + DateTime.Now.Month + "月\\"))
                Directory.CreateDirectory(Application.StartupPath + "\\" + DateTime.Now.Month + "月\\");
            if (!Directory.Exists(Application.StartupPath + "\\" + DateTime.Now.Month + "月\\" + "OCV\\"))
                Directory.CreateDirectory(Application.StartupPath + "\\" + DateTime.Now.Month + "月\\" + "OCV\\");
            StreamWriter sw = new StreamWriter(Application.StartupPath + "\\" + DateTime.Now.Month + "月\\" + "OCV\\" + DateTime.Now.ToString("yyyyMMdd") + ".txt", true, Encoding.Default);
            sw.WriteLine(msg);
            sw.Flush();
            sw.Close();
        }

        private static bool FindMsgIndex(string msg)
        {
            return msg.Contains(strCurMsg);
        }

        public void RpjLive()//记录线程启动时长以及最近一次时间，每三秒加一
        {
            while (true)
            {
                iLive += 1;
                string strTable = "sys_live_table";
                string cmd = string.Format("update {0} set sys_live='{1}',sys_time=now() where sys_name='ocv'", strTable, iLive);
                mSql.ExecuteSql(cmd);
                Thread.Sleep(3000);

            }
        }

        public void write()
        {
            iniFile.WriteValue("sys", "putstep", PutStep);
        }

        private void xieshuju(int path)//写入数据
        {
            iniFile.WriteValue("SHUJU", "tray", tbTray.Text);
            iniFile.WriteValue("SHUJU", "type", textBox5.Text);
            //for (int i = 0; i < 24; i++)
            //{
            iniFile.WriteValue("SHUJU", "lx" + path, lb[path - 1].BackColor == Color.Green ? "0" : lb[path - 1].BackColor == Color.Red ? "1" : "2");
            iniFile.WriteValue("SHUJU", "bx" + path, tb[path - 1].BackColor == Color.Green ? "1" : tb[path - 1].BackColor == Color.Red ? "2" : "0");
            iniFile.WriteValue("SHUJU", "batty" + path + "", tb[path - 1].Text);
            iniFile.WriteValue("CESHI", "v" + path.ToString(), dianya[path - 1].ToString());
            iniFile.WriteValue("CESHI", "r" + path.ToString(), naizu[path - 1].ToString());
            iniFile.WriteValue("CESHI", "t" + path.ToString(), wendu[path - 1].ToString());
            iniFile.WriteValue("CESHI", "jieguo" + path.ToString(), jg[path - 1].ToString());
            iniFile.WriteValue("CESHI", "shellVol" + path.ToString(), shellVol[path - 1].ToString());
            //}
        }

        private void clearAllData()
        {
            iniFile.WriteValue("SHUJU", "tray", tbTray.Text);
            iniFile.WriteValue("SHUJU", "type", textBox5.Text);
            for (int i = 0; i < 24; i++)
            {
                iniFile.WriteValue("SHUJU", "lx" + (i + 1), "2");
                iniFile.WriteValue("SHUJU", "bx" + (i + 1), "0");
                iniFile.WriteValue("SHUJU", "batty" + (i + 1) + "", "");
                iniFile.WriteValue("CESHI", "v" + (i + 1).ToString(), "0");
                iniFile.WriteValue("CESHI", "r" + (i + 1).ToString(), "0");
                iniFile.WriteValue("CESHI", "t" + (i + 1).ToString(), "0");
                iniFile.WriteValue("CESHI", "jieguo" + (i + 1).ToString(), "");
                iniFile.WriteValue("CESHI", "shellVol" + (i + 1).ToString(), "0");
            }
        }


        public void tongji()//生产统计数据，显示数据
        {
            if (textBox5.Text == "O1")
            {
                int o16 = int.Parse(tbo16.Text);
                int o11 = int.Parse(tbo11.Text);
                int o12 = int.Parse(tbo12.Text);
                int o13 = int.Parse(tbo13.Text);
                int o14 = int.Parse(tbo14.Text);
                int o15 = int.Parse(tbo15.Text);
                for (int i = 0; i < 24; i++)
                {
                    if (tb[i].BackColor != System.Drawing.SystemColors.Control)
                    {
                        o16++;
                        if (jg[i] == "电压不良") o11++;
                        if (jg[i] == "内阻不良") o12++;
                        if (jg[i] == "壳体不良") o13++;
                        if (jg[i] == "温度不良") o14++;
                        if (jg[i] == "找不到O1") o15++;
                    }
                }
                tbo11.Text = o11.ToString();
                tbo12.Text = o12.ToString();
                tbo13.Text = o13.ToString();
                tbo14.Text = o14.ToString();
                tbo15.Text = o15.ToString();
                tbo16.Text = o16.ToString();
                if (webtrue)
                {
                    UpdateOCVO1ShuJu(o11, o12, o14, o13, o16, float.Parse(tbo1yl.Text.Substring(0, (tbo1yl.Text.Length - 1))), 1);
                }

            }
            if (textBox5.Text == "OB")
            {
                int ob6 = int.Parse(tbob6.Text);
                int ob1 = int.Parse(tbob1.Text);
                int ob2 = int.Parse(tbob2.Text);
                int ob3 = int.Parse(tbob3.Text);
                int ob4 = int.Parse(tbob4.Text);
                int ob5 = int.Parse(tbob5.Text);
                for (int i = 0; i < 24; i++)
                {
                    if (tb[i].BackColor != System.Drawing.SystemColors.Control)
                    {
                        ob6++;
                        if (jg[i] == "电压不良") ob1++;
                        if (jg[i] == "内阻不良") ob2++;
                        if (jg[i] == "壳体不良") ob3++;
                        if (jg[i] == "温度不良") ob4++;
                        if (jg[i] == "找不到OB") ob5++;
                    }
                }
                tbob1.Text = ob1.ToString();
                tbob2.Text = ob2.ToString();
                tbob3.Text = ob3.ToString();
                tbob4.Text = ob4.ToString();
                tbob5.Text = ob5.ToString();
                tbob6.Text = ob6.ToString();
                if (webtrue)
                {
                    UpdateOCVO1ShuJu(ob1, ob2, ob4, ob3, ob6, float.Parse(tbobyl.Text.Substring(0, (tbobyl.Text.Length - 1))), 2);
                }
            }



        }


        private void xietongji()//统计写入配置文件
        {
            for (int i = 0; i < 6; i++)
            {
                iniFile.WriteValue("TONGJI", ("O1" + i).ToString(), o1[i].Text);
                iniFile.WriteValue("TONGJI", ("OB" + i).ToString(), ob[i].Text);
            }
            iniFile.WriteValue("TONGJI", "shijian", textBox1.Text);

        }

        public void saveMESLog(string LogValus)
        {
            StreamWriter sw = null;
            try
            {
                if (!Directory.Exists("D:\\MESLog"))
                {
                    Directory.CreateDirectory("D:\\MESLog");
                }
                string mpth = "D:\\MESLog" + "\\" + DateTime.Now.ToString("yyyyMMdd");
                if (!Directory.Exists(mpth))
                    Directory.CreateDirectory(mpth);

                sw = File.AppendText(mpth + "\\" + "MES测试日志.txt");
                sw.WriteLine("时间:" + DateTime.Now.ToString());
                sw.WriteLine("信息：" + LogValus + "\r\n");
                sw.Close();
            }
            catch (Exception)
            {

                if (sw != null)
                {
                    sw.Close();
                    sw.Dispose();
                }
            }
            //if (!Directory.Exists("D:\\MESLog"))
            //{
            //    Directory.CreateDirectory("D:\\MESLog");
            //}
            //string pathMES = "D:\\MESLog" + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + "MES日志.txt";
            //if (!File.Exists(pathMES))
            //{
            //    File.CreateText(pathMES);
            //}
            ////FileStream fs1 = new FileStream("D:\\MESLog" + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + "MES日志.txt", FileMode.Create, FileAccess.Write);//创建写入文件 
            //StreamWriter sw = new StreamWriter(pathMES, true, Encoding.UTF8);
            //sw.WriteLine(LogValus);//开始写入值
            //sw.Flush();
            //sw.Close();
        }

        public void baocun()//电压值,内阻值,温度值未知，须从三张表中获取，保存为CSV文件
        {
            if (!Directory.Exists("D:\\OCV\\"))
                Directory.CreateDirectory("D:\\OCV\\");
            string weizhi = "D:\\OCV\\" + DateTime.Now.ToString("yyyy-MM-dd") + "测试数据.csv";
            if (!File.Exists(weizhi))
            {
                StreamWriter sww = new StreamWriter(weizhi, true, Encoding.Default);
                sww.WriteLine("托盘条码,电池条码,测试类型,通道号,电压值,内阻值,温度值,壳电压值,测试结果,测试时间");
                sww.Close();
            }
            int iCount = 0;
            StreamWriter sw = null;
            bool blMasterSaveOK = false;
            do
            {
                try
                {
                    if (iCount == 0)
                        sw = new StreamWriter(weizhi, true, Encoding.Default);
                    else
                        sw = new StreamWriter("D:\\OCV\\测试数据" + iCount + ".csv", true, Encoding.Default);
                    blMasterSaveOK = true;
                    Thread.Sleep(200);
                }
                catch (IOException iEx)
                {
                    ShowErrMsg(iEx.Message);
                    iCount += 1;
                    if (iCount > 3)
                    {
                        foreach (System.Diagnostics.Process p in System.Diagnostics.Process.GetProcesses())
                        {
                            if (p.ProcessName.ToUpper() == "EXCEL")
                                p.Kill();
                        }
                        iCount = 0;
                    }
                }
            }
            while (!blMasterSaveOK);
            for (int i = 0; i < 24; i++)
            {
                if (tb[i].Text != "")
                {
                    sw.WriteLine(tbTray.Text + "," + tb[i].Text + "," + textBox5.Text + "," + (i + 1).ToString() + "," + dianya[i] + "," + naizu[i] + "," + wendu[i] + "," + shellVol[i] + "," + jg[i] + "," + DateTime.Now.ToString("HH:mm:ss"));
                }
            }
            sw.Close();
        }
        #endregion


        #region 条件重设
        private void btnReset_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("确认重设条件？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1) == DialogResult.Yes)
            {
                btnAccept.Enabled = true;
                groupBox5.Enabled = true;
            }
        }
        #endregion


        #region 应用条件
        private void btnAccept_Click(object sender, EventArgs e)
        {
            if (txtO1A.Text == "" || txtO1B.Text == "")
            {
                lbmessage.Text = "范围值不能为空！";
                txtO1A.Text = "";
                txtO1B.Text = "";
            }
            else
            {
                try
                {
                    if (float.Parse(txtO1A.Text) >= float.Parse(txtO1B.Text))
                    {
                        lbmessage.Text = "范围值设置有误，请重新设置！";
                        txtO1A.Text = "";
                        txtO1B.Text = "";
                    }
                    else
                    {
                        lbmessage.Text = "";
                        //txtV2.Enabled = false;
                        //txtV1.Enabled = false;
                        iniFile.WriteValue("TIAOJIAN", "MINO1VOL", txtO1A.Text);
                        iniFile.WriteValue("TIAOJIAN", "MAXO1VOL", txtO1B.Text);
                    }
                }
                catch
                {
                    lbmessage.Text = "请输入正确的数字类型！";
                    txtO1A.Text = "";
                    txtO1B.Text = "";
                }
            }
            if (txtOBA.Text == "" || txtOBB.Text == "")
            {
                lbmessage.Text = "范围值不能为空！";
                txtOBA.Text = "";
                txtOBB.Text = "";
            }
            else
            {
                try
                {
                    if (float.Parse(txtOBA.Text) >= float.Parse(txtOBB.Text))
                    {
                        lbmessage.Text = "范围值设置有误，请重新设置！";
                        txtOBA.Text = "";
                        txtOBB.Text = "";
                    }
                    else
                    {
                        lbmessage.Text = "";
                        //txtV2.Enabled = false;
                        //txtV1.Enabled = false;
                        iniFile.WriteValue("TIAOJIAN", "MINOBVOL", txtOBA.Text);
                        iniFile.WriteValue("TIAOJIAN", "MAXOBVOL", txtOBB.Text);
                    }
                }
                catch
                {
                    lbmessage.Text = "请输入正确的数字类型！";
                    txtOBA.Text = "";
                    txtOBB.Text = "";
                }
            }
            if (txtR1.Text == "" || txtR2.Text == "")
            {
                lbmessage.Text = "范围值不能为空！";
                txtR1.Text = "";
                txtR2.Text = "";
            }
            else
            {
                try
                {
                    if (float.Parse(txtR1.Text) >= float.Parse(txtR2.Text))
                    {
                        lbmessage.Text = "范围值设置有误，请重新设置！";
                        txtR1.Text = "";
                        txtR2.Text = "";
                    }
                    else
                    {
                        lbmessage.Text = "";
                        //txtR2.Enabled = false;
                        //txtR1.Enabled = false;
                        iniFile.WriteValue("TIAOJIAN", "MINR", txtR1.Text);
                        iniFile.WriteValue("TIAOJIAN", "MAXR", txtR2.Text);
                    }
                }
                catch
                {
                    lbmessage.Text = "请输入正确的数字类型！";
                    txtR1.Text = "";
                    txtR2.Text = "";
                }
            }
            if (txtShell1.Text == "" || txtShell2.Text == "")
            {
                lbmessage.Text = "范围值不能为空！";
                txtShell1.Text = "";
                txtShell2.Text = "";
            }
            else
            {
                try
                {
                    if (float.Parse(txtShell1.Text) >= float.Parse(txtShell2.Text))
                    {
                        lbmessage.Text = "范围值设置有误，请重新设置！";
                        txtShell1.Text = "";
                        txtShell2.Text = "";
                    }
                    else
                    {
                        lbmessage.Text = "";
                        //txtShell2.Enabled = false;
                        //txtShell1.Enabled = false;
                        iniFile.WriteValue("TIAOJIAN", "MINSHELL", txtShell1.Text);
                        iniFile.WriteValue("TIAOJIAN", "MAXSHELL", txtShell2.Text);
                    }
                }
                catch
                {
                    lbmessage.Text = "请输入正确的数字类型！";
                    txtShell1.Text = "";
                    txtShell2.Text = "";
                }
            }
            if (txbT1.Text == "" || txbT2.Text == "")
            {
                lbmessage.Text = "范围值不能为空！";
                txbT1.Text = "";
                txbT2.Text = "";
            }
            else
            {
                try
                {
                    if (float.Parse(txbT1.Text) >= float.Parse(txbT2.Text))
                    {
                        lbmessage.Text = "范围值设置有误，请重新设置！";
                        txbT1.Text = "";
                        txbT2.Text = "";
                    }
                    else
                    {
                        lbmessage.Text = "";
                        //txtV2.Enabled = false;
                        //txtV1.Enabled = false;
                        iniFile.WriteValue("TIAOJIAN", "MINT", txbT1.Text);
                        iniFile.WriteValue("TIAOJIAN", "MAXT", txbT2.Text);
                    }
                }
                catch
                {
                    lbmessage.Text = "请输入正确的数字类型！";
                    txbT1.Text = "";
                    txbT2.Text = "";
                }
            }
            if (txbNG.Text == "")
            {
                lbmessage.Text = "范围类型不能为空！";
            }
            else
            {
                try
                {
                    if (int.Parse(txbNG.Text) > 24 || int.Parse(txbNG.Text) < 0)
                    {
                        txbNG.Text = "";
                        lbmessage.Text = "请输入正确的NG个数！0-24";
                    }
                    else
                        iniFile.WriteValue("TIAOJIAN", "NGNUM", txbNG.Text);
                }
                catch
                {
                    lbmessage.Text = "请输入正确的数字类型!";
                    txbNG.Text = "";
                }

            }

            if (lbmessage.Text == "")
            {
                groupBox5.Enabled = false;
                btnAccept.Enabled = false;
            }
        }
        #endregion


        #region 补码
        private void btnFill_Click(object sender, EventArgs e)
        {
            if (txbFill.Text != "")
            {
                trayCode = txbFill.Text;
                txbFill.ReadOnly = true;
                ShowErrMsg("补码成功");
            }
            if (comboBox1.Text != "")
            {
                type = comboBox1.Text == "O1" ? 1 : comboBox1.Text == "OB" ? 3 : 2;

            }
            //string fillSql = string.Format("insert into ocv_test_type(tray_code) values('{0}')", cgVar.mLoad.sTraycode);
            //int fRes = mSql.ExecuteSql(fillSql);
            //if (fRes > 0 )
            //{
            //    ShowErrMsg("补码成功
        }
        #endregion


        #region 启动程序
        private void button2_Click(object sender, EventArgs e)
        {
            if (button2.Text == "运行")
            {
                //string[] cell = { "123B11111"};
                //for (int i = 0; i < 24; i++)
                //{
                //    string sql = string.Format("insert into ocvdisplay(id,TrayCode,BatteryCode,Passagewary,BatteryState) values({0},'{1}','{2}',{3},{4})", (i + 1), "6F100001", cell[0], (i + 1), 1);
                //    mysql3.ExecuteSql(sql);
                //}
                NGCount = 0;
                for (int i = 0; i < 24; i++)
                {
                    if (tb[i].BackColor == Color.Red)
                        NGCount++;
                }
                label22.Text = "当前NG个数:" + NGCount;
                try
                {
                    if (string.IsNullOrEmpty(strIP2) || iPort2 == 0)//strIP2是PLC默认IP
                    {
                        ShowErrMsg("PLC通信失败，请检查网络通信地址端口等是否正常！");
                        return;
                    }

                    if (plcNetOperator == null)
                        plcNetOperator = new clientTest(strIP2, iPort2);//"192.168.10.202", 9094);

                    if (!plcNetOperator.isConnected)
                        plcNetOperator.connectServer();

                    Thread.Sleep(1000);

                    if (!plcNetOperator.isConnected)
                    {
                        ShowErrMsg("PLC通信失败，请检查网络接口是否正常！");
                        return;
                    }
                    else
                    {
                        label18.BackColor = Color.Green;
                    }

                    if (!sp3562.IsOpen)//内阻表串口连接

                        sp3562.Open();//sp3562是一个IO对象
                    gettemperature_normal();//温度表串口连接

                    sp3562.ReadExisting();
                    sp3562.WriteLine("*IDN?");

                    Thread.Sleep(100);
                    string bm = sp3562.ReadExisting();
                    if (bm == "")
                    {
                        ShowErrMsg("与3562表通讯异常");
                        return;
                    }

                    sp3562.WriteLine(":RESISTANCE:RANGE 3.0000E-3");
                    if (ctTest1 == null)//电压表网口连接
                    {
                        ctTest1 = new clientTest(strIP, iPort);//clientTest是一个连接其他设备的专用类
                        ctTest1.connectServer();
                    }
                    if (!ctTest1.isConnected)
                    {
                        ShowErrMsg("与34461表通讯异常");
                        return;
                    }
                    //measure();
                    if (fnormal_temp == 0)
                    {
                        ShowErrMsg("与测温仪1通讯异常");
                        return;
                    }
                    if (fnormal_temp2 == 0)
                    {
                        ShowErrMsg("与测温仪2通讯异常");
                        return;
                    }
                    plcComm.plc = plcNetOperator;
                    if (plcComm.SetDevice("R0332", 1) > 0)
                    {
                        if (thOCVLive == null)
                        {
                            thOCVLive = new Thread(new ThreadStart(this.RpjLive));
                            thOCVLive.IsBackground = true;
                            thOCVLive.Start();
                        }

                        timerun.Enabled = true;
                        button2.Text = "停止";
                        if (webtrue)
                            UpdateOCVIsConnectPlc(1);
                    }
                }
                catch (Exception err)
                {
                    ShowErrMsg(err.Message.ToString());
                }
            }
            else
            {
                if (thOCVLive != null)
                {
                    thOCVLive.Abort();
                    thOCVLive = null;
                }
                button2.Text = "运行";
                timerun.Enabled = false;
                //cgVar.mLoad.portClose();//串口关闭
                UpdateOCVIsConnectPlc(0);

            }
        }
        #endregion



        #region 获取温度/获取电压内阻/检测校准时间
        float fnormal_temp = 0;
        float fnormal_temp2 = 0;
        private void gettemperature_normal()//获取温度
        {
            try
            {
                fnormal_temp = 0;
                fnormal_temp2 = 0;
                if (!sptemp.IsOpen)//sptemp是一个IOPort对象
                    sptemp.Open();
                sptemp.ReadExisting();
                byte[] data = { 0x1, 0x3, 0x0, 0x0, 0x0, 0x2, 0xC4, 0xB };
                try
                {
                    sptemp.Write(data, 0, data.Length);
                    Thread.Sleep(100);
                    sptemp.Read(data, 0, 7);
                }
                catch (Exception er)
                {
                    ShowErrMsg(er.Message);
                }
                if (data[2] == 0x0)
                {
                    fnormal_temp = 0;
                }
                else
                {
                    fnormal_temp = (Convert.ToSingle((data[5] * 256 + data[6])) / 10);
                }
                if (!sptemp2.IsOpen)//sptemp是一个IOPort对象
                    sptemp2.Open();
                sptemp.ReadExisting();
                byte[] data1 = { 0x1, 0x3, 0x0, 0x0, 0x0, 0x2, 0xC4, 0xB };
                try
                {
                    sptemp.Write(data1, 0, data.Length);
                    Thread.Sleep(100);
                    sptemp.Read(data1, 0, 7);
                }
                catch (Exception er)
                {
                    ShowErrMsg(er.Message);
                }
                if (data1[2] == 0x0)
                {
                    fnormal_temp2 = 0;
                }
                else
                {
                    fnormal_temp2 = (Convert.ToSingle((data[5] * 256 + data[6])) / 10);
                }
            }
            catch
            {
                fnormal_temp = 0;
                fnormal_temp2 = 0;
            }
        }

        public void measure()//得到电压电阻
        {
            sp3562.ReadExisting();
            //ctTest1.sendString("READ?\r");
            //ctTest1 = new clientTest(strIP, iPort);
            
            ctTest1.sendString("MEAS:VOLT:DC?\r\n");//发送读取电压表的指令
            sp3562.WriteLine(":FETCH?");//发送读取内阻表的指令
            Thread.Sleep(500);
            string sRec1 = ctTest1.recString();
            try
            {
                string[] sArray1 = sRec1.Split('\r', '\n');
                v = System.Math.Abs(Convert.ToSingle(sArray1[0]) + Convert.ToSingle(tbvd1.Text));
            }
            catch
            {
                v = -1;
            }
            try
            {
                string[] tempv = sp3562.ReadExisting().Split(',');
                r = Convert.ToSingle(tempv[0]) * 1000 + Convert.ToSingle(tbrd1.Text);

            }
            catch
            {
                r = -1;
            }
            //MessageBox.Show("V:"+v.ToString()+"　　R:"+r.ToString());
        }


        private bool GetClearBT3562()//检测校准时间是否已经到
        {
            if (string.IsNullOrEmpty(nextAdjDateTime))
            {
                return true;
            }

            DateTime nextAdjDatetime = Convert.ToDateTime(nextAdjDateTime);
            if (DateTime.Now >= nextAdjDatetime)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private void MasterExam()
        {
            
        }
        
        #endregion



        #region 时间控件 核芯代码
        private void timerun_Tick(object sender, EventArgs e)
        {


            #region Master校准
            if (chkMaster.Checked)
            {
                if (GetClearBT3562())
                {
                    if (isClear == 0)
                    {
                        if (plcComm.SetDevice("R0270", 1) > 0)
                        {
                            ShowErrMsg("准备清零……");
                            if (plcComm.GetDevice("R0271"))//仪表清零到位信号
                            {
                                ShowErrMsg("仪表清零开始");
                                // Thread.Sleep(10000);
                                sp3562.WriteLine(":ADJ?");
                                Thread.Sleep(1000);
                                string sRtn = sp3562.ReadExisting();
                                tbr.Text = sRtn;
                                if (sRtn.Trim() == "0")
                                {
                                    plcComm.SetDevice("R0272", 1);//仪表清零校准成功
                                    ShowErrMsg("清零完成");
                                    isClear = 1;
                                    needleCount++;
                                    label24.Text = "探针下压次数：" + needleCount;
                                }
                            }
                        }
                    }
                    if (plcComm.GetDevice("R0274"))//plc发送内阻校正信号
                    {
                        try
                        {
                            ShowErrMsg("内阻块校准启动.");

                            /*可能要加信号点监测*/

                            if (!sp3562.IsOpen)//3562校准串口
                                sp3562.Open();
                            sp3562.ReadExisting();
                            sp3562.WriteLine(":RESISTANCE:RANGE 3.0000E-3");
                            sp3562.WriteLine(":FETC?");
                            Thread.Sleep(1000);
                            Single sngMaster = Single.Parse(sp3562.ReadExisting().Trim().Split(',')[0]) * 1000;
                            tbr.Text = sngMaster.ToString();
                            Single sngMasterStander = Single.Parse(tbMaster.Text.ToString());
                            Single sngMasterOffer = Single.Parse(tbMasterOffer.Text.ToString());
                            master = sngMaster - sngMasterStander;
                            masterSum += master;
                            tbrd1.Text = iniFile.GetString("XITONG", "masterChazhi", "");
                            string strResult = Math.Abs(master) > sngMasterOffer ? "NG" : "OK";
                            if (!Directory.Exists("D:\\Master\\"))
                                Directory.CreateDirectory("D:\\Master\\");
                            if (!File.Exists("D:\\Master\\master.csv"))
                            {
                                StreamWriter sww = new StreamWriter("D:\\Master\\master.csv", true, Encoding.Default);
                                sww.WriteLine("标准值,测量值,差值,判断标准,结果,测量时间");
                                sww.Close();
                            }
                            int iCount = 0;
                            StreamWriter sw = null;
                            bool blMasterSaveOK = false;
                            do
                            {
                                try
                                {
                                    if (iCount == 0)
                                        sw = new StreamWriter("D:\\Master\\master.csv", true, Encoding.Default);
                                    else
                                        sw = new StreamWriter("D:\\Master\\master_" + iCount + ".csv", true, Encoding.Default);
                                    blMasterSaveOK = true;
                                    Thread.Sleep(200);
                                }
                                catch (IOException iEx)
                                {
                                    ShowErrMsg(iEx.Message);
                                    iCount += 1;
                                    if (iCount > 3)
                                    {
                                        foreach (System.Diagnostics.Process p in System.Diagnostics.Process.GetProcesses())
                                        {
                                            if (p.ProcessName.ToUpper() == "EXCEL")
                                                p.Kill();
                                        }
                                        iCount = 0;
                                    }
                                }
                            }
                            while (!blMasterSaveOK);
                            sw.WriteLine(tbMaster.Text + "," + sngMaster + "," + master + "," + sngMasterOffer + "," + strResult + "," + DateTime.Now.ToString());
                            sw.Close();
                            float fixResult = Math.Abs(master);
                            if (fixResult > sngMasterOffer)
                            {
                                plcComm.SetDevice("R0277", 1);//master测试NG
                                needleCount++;
                                label24.Text = "探针下压次数：" + needleCount;
                            }
                            else
                            {
                                if (plcComm.SetDevice("R0275", 1) > 0)
                                {
                                    needleCount++;
                                    adjCount++;
                                    label24.Text = "探针下压次数：" + needleCount;
                                    if (adjCount == 10)
                                        iniFile.WriteValue("XITONG", "masterChazhi", (masterSum / 10));
                                    else
                                        ShowErrMsg("第" + adjCount + "次校准成功！");
                                }//通知PLC校准成功
                                //plcComm.SetDevice("R0275", 1);//上位机通知PLC校准结束

                                //isClear = 0;
                            }
                        }
                        catch (Exception err)
                        {
                            ShowErrMsg(err.Message.ToString());
                        }
                    }
                    if (plcComm.GetDevice("R0278"))
                    {
                        if (DateTime.Now.Hour < Convert.ToDateTime(strStrartIime).Hour)
                        {
                            nextAdjDateTime = DateTime.Now.Add(Convert.ToDateTime(strStrartIime) - DateTime.Now).ToString();
                        }
                        else
                        {
                            TimeSpan iAddAdjHoure = DateTime.Now - Convert.ToDateTime(strStrartIime);
                            TimeSpan dt = TimeSpan.FromHours((double)iAdjOffHour - (iAddAdjHoure.TotalHours % iAdjOffHour));
                            nextAdjDateTime = DateTime.Now.Add(dt).ToString();
                        }
                        //iniFile.WriteValue("CLEAR", "starttime", nextAdjDateTime);
                        iniFile.WriteValue("CLEAR", "time", nextAdjDateTime);
                        iniFile.WriteValue("CLEAR", "isjz", "1");
                        tbrd1.Text = iniFile.GetString("XITONG", "masterChazhi", "");
                        plcComm.SetDevice("R0278", 0);
                    }
                    
                }
            }
            else
            {
                //ShowErrMsg("未启动Master校准！");
                //return;
                //plcComm.SetDevice("R004A", 0);//上位机通知PLC校准结束
            }
            #endregion



            if (!plcComm.GetDevice("R0233"))//自动状态
            {
                ShowErrMsg("设备不在自动状态！");
                //button2.Text = "运行";
                return;
            }

            if (!plcNetOperator.isConnected)
            {
                if (plcNetOperator != null)
                    plcNetOperator = null;

                plcNetOperator = new clientTest(strIP2, iPort2);//"192.168.10.150", 9094);
                plcNetOperator.connectServer();//创建PLC服务连接
                Thread.Sleep(1000);
                ShowErrMsg("timer运行时，正在连接PLC...");
                if (webtrue)
                {
                    UpdateOCVIsConnectPlc(1);
                }
                if (plcNetOperator.isConnected)
                {
                    if (webtrue)
                    {
                        UpdateOCVIsConnectPlc(0);
                    }
                    ShowErrMsg("已重新创建PLC连接");
                }
                plcComm.plc = plcNetOperator;
                return;
            }





            #region  取放盘


            #endregion


            #region 获取调度请求和层状态

            #endregion


            #region OCV取放盘

            # endregion


            #region 测试
            string strRtn2 = "";
            string[] pts2 = { "R0255", "R0290", "R5046", "R0292" };
            bool blShowStatus2 = plcComm.ReadMoreDevice(out strRtn2, pts2) == 0 ? true : false;
            if (blShowStatus2)
            {
                int scan = int.Parse(strRtn2.Substring(0, 1));//R0021扫码信号（获取测试类型和电池数据）
                int ceshi = int.Parse(strRtn2.Substring(1, 1));//R0023测试开始信号
                int end = int.Parse(strRtn2.Substring(2, 1));//R0026总测试完成信号
                int shell = int.Parse(strRtn2.Substring(3, 1));
                if (scan == 1)
                {
                    string cellCode = "";
                    string mOut = "";
                    int sCount = 0;
                    if (isScanned == 0)
                    {
                        ShowErrMsg("OCV请求扫码");
                        try
                        {
                            //cgVar.mLoad.explainStatusRead();//状态读取说明方法
                            trayCode = cgVar.mLoad.getCodeBySerialData(scannerConn);//开始托盘扫码
                            //Thread.Sleep(500);
                            if (trayCode.Length > 0)
                            {
                                if (trayCode.Substring(0, 5) == "error")
                                {
                                    while (sCount < 10)
                                    {
                                        trayCode = cgVar.mLoad.getCodeBySerialData(scannerConn);
                                        if (trayCode.Length > 0 && trayCode.Substring(0, 5) != "error")
                                        {
                                            break;
                                        }
                                        sCount++;
                                    }
                                    if (sCount <= 9)
                                    {
                                        ShowErrMsg("OCV扫码成功");
                                        //plcComm.SetDevice(cgVar.mLoad.sPlcTrayScanOk, 1);
                                    }
                                    else
                                    {
                                        ShowErrMsg("扫码失败，请手动输入条码");
                                        txbFill.ReadOnly = false;
                                        isScanned = 1;
                                        return;
                                    }

                                }
                            }
                            else
                            {
                                while (sCount < 10)
                                {
                                    trayCode = cgVar.mLoad.getCodeBySerialData(scannerConn);
                                    if (trayCode.Length > 0 && trayCode.Substring(0, 5) != "error")
                                    {
                                        break;
                                    }
                                    sCount++;
                                }
                                if (sCount <= 9)
                                {
                                    ShowErrMsg("OCV扫码成功");
                                    //plcComm.SetDevice(cgVar.mLoad.sPlcTrayScanOk, 1);
                                }
                                else
                                {
                                    ShowErrMsg("扫码失败，请手动输入条码");
                                    txbFill.ReadOnly = false;
                                }

                            }
                        }
                        catch
                        {
                            ShowErrMsg("扫码异常，请检查扫码枪状态！");
                            return;
                        }
                        isScanned = 1;
                    }
                    string[] tc = trayCode.Split('\r', '\n');
                    //trayCode = tc[0];
                    tbTray.Text = tc[0];
                    int channel = 0;
                    if (tbTray.Text != "")
                    {
                        try
                        {
                            string findCellSql = string.Format("select battery_code,channel_no from tray_onbind_battery where tray_code='{0}'", tbTray.Text);
                            DataTable dt = mSql.GetDataTableByCmd(findCellSql, out mOut);
                            rowsCount = dt.Rows.Count;
                            if (rowsCount > 0)
                            {
                                for (int i = 0; i < rowsCount; i++)
                                {
                                    channel = int.Parse(dt.Rows[i]["channel_no"].ToString());
                                    cellCode = dt.Rows[i]["battery_code"].ToString();
                                    tb[channel - 1].Text = cellCode;
                                }
                                for (int j = 0; j < 24; j++)
                                {
                                    if (tb[j].Text == "")
                                    {
                                        ShowErrMsg("未从绑定关系表中获取到" + j + "通道电芯条码");
                                        tb[j].BackColor = Color.Red;
                                    }
                                }
                            }
                            else
                            {
                                ShowErrMsg("条码在数据库中没有匹配，请检查或补码");
                                txbFill.ReadOnly = false;
                                return;
                            }
                        }
                        catch
                        {
                            ShowErrMsg("数据库连接异常，请检查");
                            return;
                        }
                        for (int j = 0; j < 24; j++)
                        {
                            iniFile.WriteValue("SHUJU", "batty" + (j + 1), tb[j].Text);
                        }
                        string strSend = "", strSend1 = "";//电池一列和电池二列，从右向左
                        for (int i = 0; i < tb.Length; i++)
                        {
                            if (i < 12)
                            {
                                strSend = (string.IsNullOrEmpty(tb[i].Text) == true ? "0" : "1") + strSend;//电池一列,1为有电池,0无
                            }
                            else
                            {
                                strSend1 = (string.IsNullOrEmpty(tb[i].Text) == true ? "0" : "1") + strSend1;//电池二列,1为有电池,0无
                            }
                        }
                        int[] mdata1 = new int[1];
                        int[] mdata2 = new int[1];
                        mdata1[0] = Convert.ToInt32(strSend, 2);
                        mdata2[0] = Convert.ToInt32(strSend1, 2);
                        string sOut2 = "";
                        string dateSql = string.Format("select datetime from tray_onbind_battery where tray_code='{0}'", tbTray.Text);
                        TimeSpan standByTime;
                        DataTable dTable = mSql.GetDataTableByCmd(dateSql, out sOut2);
                        if (dTable.Rows.Count > 0)
                        {
                            DateTime date = Convert.ToDateTime(dTable.Rows[0]["datetime"].ToString());
                            if (date != null)
                            {
                                standByTime = DateTime.Now - date;
                                timeRes = standByTime.TotalHours;
                            }
                            else
                            {
                                ShowErrMsg("未能从绑定关系表中获取时间信息，请手输类型");
                                return;
                            }
                        }
                        if (isWriteType == 0)
                        {
                            string sOut = "";

                            try
                            {
                                string typeSql = string.Format("select load_status from ocv_test_type where tray_code='{0}'", tbTray.Text);
                                //string typeSql = string.Format("select load_status from ocv_test_type where tray_code='{0}'", cgVar.mLoad.sTraycode);
                                DataTable tempTable = mSql.GetDataTableByCmd(typeSql, out sOut);
                                if (tempTable.Rows.Count > 0)
                                {
                                    type = int.Parse(tempTable.Rows[0]["load_status"].ToString());

                                }
                                else
                                {
                                    ShowErrMsg("未能获取到对应托盘条码的数据类型");
                                    ShowErrMsg("正根据静置时长自动补充……");
                                    //textBox2.ReadOnly = false;

                                    if (timeRes >= 15 && timeRes < 45)
                                    {
                                        type = 1;
                                    }
                                    else if (timeRes >= 45)
                                    {
                                        type = 3;
                                    }
                                    else
                                    {
                                        type = 4;
                                        ShowErrMsg("托盘" + tbTray.Text + "静置时间未到，即将回流");
                                    }
                                }
                                isWriteType = 1;
                            }
                            catch
                            {
                                ShowErrMsg("数据库异常，请检查参数值");
                                return;
                            }
                        }
                        //测试类型防呆
                        if (isFangdai == 0)
                        {
                            if (timeRes < 14)
                            {    
                                if (type == 1)
                                {
                                    timerun.Enabled = false;
                                    if (MessageBox.Show("静置时间未达到15小时，是否进行O1测试？", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.No)
                                    {
                                        //给plc发信号
                                        type = 4;
                                        ShowErrMsg("托盘回流");
                                        timerun.Enabled = true;
                                        isFangdai = 1;
                                    }
                                }
                                else
                                {
                                    ShowErrMsg("静置时间未到，不能进行OB测试，即将回流！");
                                    type = 4;
                                    timerun.Enabled = true;
                                    isFangdai = 1;
                                }
                            }
                            else if (timeRes > 15 && timeRes < 45)
                            {
                                
                                if (type == 3)
                                {
                                    timerun.Enabled = false;
                                    if (MessageBox.Show("静置时间未达到45小时，是否进行OB测试？", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.No)
                                    {
                                        //给PLC发信号
                                        type = 1;
                                        ShowErrMsg("已更改类型为O1");
                                        timerun.Enabled = true;
                                        isFangdai = 1;
                                    }
                                }
                            }
                            else
                            {
                                
                                if (type == 1)
                                {
                                    timerun.Enabled = false;
                                    if (MessageBox.Show("静置时间达到或超过45小时，是否进行O1测试？", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.No)
                                    {
                                        type = 3;
                                        ShowErrMsg("已更改类型为OB");
                                        timerun.Enabled = true;
                                        isFangdai = 1;
                                    }
                                }
                            }
                        }
                        if (type != -1)
                        {
                            textBox5.Text = type == 1 ? "O1" : type == 3 ? "OB" : "Error";
                            //textBox5.Text = textBox2.Text;
                            if (textBox5.Text == "Error")
                            {
                                textBox5.BackColor = Color.Red;
                                ShowErrMsg("该盘静置时间有误，即将排出！");
                            }
                            int[] ldata = new int[1];
                            if (type == 4)
                            {
                                type = 3;
                            }
                            ldata[0] = type;
                            textBox5.BackColor = System.Drawing.SystemColors.Control;
                            if (plcComm.WriteDeviceBlock("32496", 1, mdata1) == 0 && plcComm.WriteDeviceBlock("32497", 1, mdata2) == 0)
                            {
                                if (plcComm.WriteDeviceBlock("32500", 1, ldata) == 0) //32500
                                {
                                    ShowErrMsg("写测试类型" + ldata[0].ToString() + "到PLC成功");
                                }
                                else
                                {
                                    ShowErrMsg("写测试类型失败！");
                                    return;
                                }
                            }
                            else
                            {
                                ShowErrMsg("写电池有无数据失败！");
                                return;
                            }

                            if (plcComm.SetDevice(cgVar.mLoad.sPlcTrayScanOk, 1) > 0)
                            {
                                ShowErrMsg("已发送扫码成功信号！");
                            }
                            //trayCode = cgVar.mLoad.sTraycode;
                            if (tbTray.Text != "")
                            {
                                trayCode = "";
                            }
                            string delSql = string.Format("delete from ocv_test_type where tray_code='{0}' ", tbTray.Text);
                            if (mSql.ExecuteSql(delSql) < 0)
                            {
                                ShowErrMsg("删除测试类型表失败，请检查数据库！");
                                return;
                            }
                            PutStep++;
                            write();
                        }
                    }
                    else
                    {
                        ShowErrMsg("请手动输入条码！");
                        txbFill.ReadOnly = false;
                        return;
                    }
                }

                if (ceshi == 1)
                {
                    int sOut = 0;
                    int[] ldata = new int[1];
                    if (plcComm.ReadDeviceBlock("32502", 1, out ldata) == 0)//获取通道
                    {

                        if (ldata[0] < 1 || ldata[0] > 24)
                        {
                            ShowErrMsg("通道号溢出!");
                            return;
                        }
                    }
                    else
                    {
                        ShowErrMsg("获取通道失败！");
                        return;
                    }
                    //if (NGCount > ldata[0])
                    //{
                    //    NGCount = ldata[0];
                    //    label22.Text = "当前NG个数：" + NGCount;
                    //    for (int i = ldata[0]; i < 24; i++)
                    //    {
                    //        tb[i].BackColor = System.Drawing.SystemColors.Control;
                    //    }
                    //}
                    //Thread.Sleep(1000);
                    measure();//获取电压电阻
                    gettemperature_normal();//获取温度
                    int weizhi = ldata[0];
                    tbPlace.Text = weizhi.ToString();
                    tbBattay.Text = tb[(weizhi - 1)].Text;
                    tbr.Text = (r - master).ToString();
                    tbv.Text = v.ToString();
                    temp = (fnormal_temp + fnormal_temp2) / 2;
                    tbt.Text = temp.ToString();
                    dianya[(ldata[0] - 1)] = float.Parse(tbv.Text);
                    naizu[(ldata[0] - 1)] = float.Parse(tbr.Text);
                    wendu[(ldata[0] - 1)] = float.Parse(tbt.Text);
                    try
                    {
                        if (type == 1)
                        {
                            if (v < float.Parse(txtO1A.Text) || v > float.Parse(txtO1B.Text))
                            {
                                sOut = 1;
                            }
                        }
                        if (type == 2)
                        {
                            if (v < float.Parse(txtOBA.Text) || v > float.Parse(txtOBB.Text))
                            {
                                sOut = 1;
                            }
                        }
                        if (r < float.Parse(txtR1.Text) || r > float.Parse(txtR2.Text))
                        {
                            sOut = 2;
                        }
                        if (temp < float.Parse(txbT1.Text) || temp > float.Parse(txbT2.Text))
                        {
                            sOut = 4;
                        }
                    }
                    catch (Exception ex)
                    {
                        ShowErrMsg(ex.Message.ToString());
                    }
                    string[] sOCV = { "良品", "电压不良", "内阻不良", "壳体不良", "温度不良" };
                    if (sOut == 0)
                    {
                        lbStatus.Text = sOCV[sOut].ToString();
                        lbStatus.ForeColor = Color.Green;
                        tb[(ldata[0] - 1)].BackColor = Color.Green;
                        //ShowErrMsg(tb[(ldata[0] - 1)].Text + ":" + iRt.ToString() + sOCV[iRt].ToString());
                        jg[(ldata[0] - 1)] = sOCV[sOut].ToString();
                        if (count == ldata[0])
                        {
                            NGCount--;
                            label22.Text = "当前NG个数：" + NGCount;
                            ShowErrMsg(tb[(ldata[0] - 1)].Text + "复测:" + sOCV[sOut].ToString());
                            try
                            {
                                DateTime date = DateTime.Now;
                                AddOcvStatistic(tb[ldata[0] - 1].Text, tbTray.Text,v, r, temp, date);
                            }
                            catch
                            {
                                ShowErrMsg("写单机数据统计失败!");
                            }
                        }                 
                        plcComm.SetDevice("R0291", 1);//单个测试完成
                        needleCount++;
                        label24.Text = "探针下压次数：" + needleCount;

                    }
                    else
                    {
                        lbStatus.Text = sOCV[sOut].ToString();
                        lbStatus.ForeColor = Color.Red;
                        tb[(ldata[0] - 1)].BackColor = Color.Red;
                        jg[(ldata[0] - 1)] = sOCV[sOut].ToString();
                        ShowErrMsg(tb[(ldata[0] - 1)].Text + ":" + sOCV[sOut].ToString());
                        if (count != ldata[0])
                        {
                            NGCount++;
                        }
                        label22.Text = "当前NG品个数：" + NGCount;
                        count = ldata[0];
                        Thread.Sleep(1000);
                        plcComm.SetDevice("R0297", 1);
                        needleCount++;
                        label24.Text = "探针下压次数：" + needleCount;
                        xieshuju(ldata[0]);
                    }
                    int color = tb[(ldata[0] - 1)].BackColor == Color.Green ? 1 : tb[(ldata[0] - 1)].BackColor == Color.Red ? 2 : 3;
                    if (webtrue)
                    {
                        try
                        {
                            AddOCV(tbTray.Text, tb[ldata[0] - 1].Text, ldata[0], color, ldata[0]);
                        }
                        catch
                        {
                            ShowErrMsg("写电池数据失败!");
                        }
                    }
                    try
                    {
                        DateTime date = DateTime.Now;
                        if (type == 1)
                        {
                            AddOcvStatistic(tb[ldata[0] - 1].Text, tbTray.Text, v, r, temp, date);
                        }
                        if (type == 2)
                        {
                            AddOcvStatisticB(tb[ldata[0] - 1].Text, v, r, temp, date);
                        }
                    }
                    catch
                    {
                        ShowErrMsg("写单机数据统计失败!");
                    }
                    try
                    {
                        string sql = string.Format("update ocv_statistic set standing_time='{0}' where tray_code='{1}'", timeRes, tbTray.Text);
                        mSql.ExecuteSql(sql);
                    }
                    catch
                    {
                        ShowErrMsg("数据库连接异常，请检查！");
                        return;
                    }
                    isUpload = 0;
                }
                if (shell == 1)
                {
                    int[] ldata1 = new int[1];
                    if (plcComm.ReadDeviceBlock("32502", 1, out ldata1) == 0)//获取通道
                    {

                        if (ldata1[0] < 1 || ldata1[0] > 24)
                        {
                            ShowErrMsg("通道号溢出!");
                            return;
                        }
                        float v2 = 0;
                        //Thread.Sleep(1000);
                        ctTest1.sendString("MEAS:VOLT:DC?\r\n");
                        Thread.Sleep(500);
                        string sRec2 = ctTest1.recString();
                        try
                        {
                            string[] sArray2 = sRec2.Split('\r', '\n');
                            v2 = System.Math.Abs(Convert.ToSingle(sArray2[0]) + Convert.ToSingle(tbvd1.Text));
                            //plcNetOperator = plcComm.plc;
                            //string rtn = plcNetOperator.recString();
                            //if (!string.IsNullOrEmpty(rtn) && rtn.Split('\r')[0].Equals("%01$WC14"))
                            //{
                            //    ShowErrMsg(ldata1[0].ToString() + "位电芯测试壳电压" + tbshell.Text);
                            //}
                        }
                        catch
                        {
                            v2 = -1;
                        }
                        int sOut = 0;
                        if (v2 < float.Parse(txtShell1.Text) || v2 > float.Parse(txtShell2.Text))
                        {
                            sOut = 4;
                        }
                        string[] sOCV = { "良品", "电压不良", "内阻不良", "K值不良", "壳体不良", "温度不良", "找不到O1", "未读取到数据" };
                        if (sOut == 4)
                        {
                            lbStatus.Text = sOCV[sOut].ToString();
                            lbStatus.ForeColor = Color.Red;
                            ShowErrMsg(tb[(ldata1[0] - 1)].Text + ":" + sOCV[sOut].ToString());
                            tb[(ldata1[0] - 1)].BackColor = Color.Red;
                            jg[(ldata1[0] - 1)] = sOCV[sOut].ToString();
                            if (count2 != ldata1[0])
                            {
                                NGCount++;
                            }
                            label22.Text = "当前NG品个数：" + NGCount;
                            count2 = ldata1[0];
                            Thread.Sleep(1000);
                            plcComm.SetDevice("R029A", 1);
                            needleCount++;
                            label24.Text = "探针下压次数：" + needleCount;
                        }
                        else
                        {
                            lbStatus.Text = sOCV[sOut].ToString();
                            lbStatus.ForeColor = Color.Green;
                            tb[(ldata1[0] - 1)].BackColor = Color.Green;
                            jg[(ldata1[0] - 1)] = sOCV[sOut].ToString();
                            if (count2 == ldata1[0])
                            {
                                NGCount--;
                                ShowErrMsg(tb[(ldata1[0] - 1)].Text + "复测:" + sOCV[sOut].ToString());
                            }
                            plcComm.SetDevice("R0293", 1);
                        }
                        int color = tb[(ldata1[0] - 1)].BackColor == Color.Green ? 1 : tb[(ldata1[0] - 1)].BackColor == Color.Red ? 2 : 3;
                        if (webtrue)
                        {
                            try
                            {
                                AddOCV(tbTray.Text, tb[ldata1[0] - 1].Text, ldata1[0], color, ldata1[0]);
                            }
                            catch
                            {
                                ShowErrMsg("写电池数据失败!");
                            }
                            try
                            {
                                UpdateOCVTestCode(tbTray.Text);
                            }
                            catch
                            {
                                ShowErrMsg("写托盘条码失败！");
                            }
                            try
                            {
                                UpdateOCVTNKC(tb[ldata1[0] - 1].Text, temp, r, v, ldata1[0]);
                            }
                            catch
                            {
                                ShowErrMsg("写电池测量数据失败!");
                            }
                            try
                            {
                                UpdateOCVShellVol(v2);
                            }
                            catch
                            {
                                ShowErrMsg("壳电压更新失败！");
                            }

                        }
                        
                        try
                        {
                            if (type == 1)
                            {
                                AddOcvStatisticShell(v2, tb[ldata1[0] - 1].Text);
                            }
                            if (type == 2)
                            {
                                AddOcvStatisticShellB(v2, tb[ldata1[0] - 1].Text);
                            }
                        }
                        catch
                        {
                            ShowErrMsg("写单机数据统计失败！");
                        }
                        shellVol[(ldata1[0] - 1)] = v2;
                        tbshell.Text = v2.ToString();
                        xieshuju(ldata1[0]);
              
                    }
                }
                //if (NGCount == int.Parse(txbNG.Text))
                //{
                //    if (MessageBox.Show("NG个数达到上限，是否继续工序？", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.No)
                //    {
                //        //向PLC发送整盘排出信号
                //    }
                //}
                float res = (float)NGCount / 24;
                textBox4.Text = (res * 100).ToString() + "%";


                if (end == 1)
                {
                    int code = 0;
                    if (isUpload == 0)
                    {
                        //List<dataCollectSfcParametricData[]> dataList = new List<dataCollectSfcParametricData[]>();
                        dataCollectSfcParametricData[] dcp = new dataCollectSfcParametricData[rowsCount];
                        for (int n = 0; n < rowsCount; n++)
                        {
                            CollectData[] cCollectData = new CollectData[] {new CollectData("OCV",dianya[n].ToString()),new CollectData("IMP",naizu[n].ToString()),
                                    new CollectData("TEMPERATURE",wendu[n].ToString()),new CollectData("CHANNELID",(n+1).ToString())};
                            //dataList.Add(miDataCollectForProcessLotForEach1(cCollectData, tbTray.Text));
                            dcp[n] = miDataCollectForProcessLotForEach1(cCollectData, tb[n].Text);
                        }
                        miDataCollectForProcessLotForEach(dcp, tbTray.Text, type, out code);
                        isUpload = 1;

                        if (code == 0)
                        {
                            plcComm.SetDevice("R5048", 1);
                            ShowErrMsg("托盘:" + tbTray.Text + textBox5.Text + "电芯数据上传MES成功！");
                        }
                        else
                        {
                            if (code != 1)
                            {
                                timerun.Enabled = false;
                                if (MessageBox.Show("托盘:" + tbTray.Text + textBox5.Text + "电芯数据上传MES失败！", "警告", MessageBoxButtons.RetryCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.Cancel)
                                {
                                    //发信号
                                    plcComm.SetDevice("R5048", 1);
                                    timerun.Enabled = true;
                                }
                                else
                                {
                                    timerun.Enabled = true;
                                    int i = 0;
                                    while (i < 3)
                                    {
                                        miDataCollectForProcessLotForEach(dcp, tbTray.Text, type, out code);
                                        if (code == 0)
                                        {
                                            break;
                                        }
                                        i++;
                                    }
                                    if (i <= 2)
                                    {
                                        ShowErrMsg("托盘:" + tbTray.Text + "电芯数据上传MES成功！");
                                        plcComm.SetDevice("R5048", 1);
                                    }
                                    else
                                        ShowErrMsg("托盘:" + tbTray.Text + "电芯数据上传MES失败！请人工处理！");
                                }
                            }
                            else
                            {
                                plcComm.SetDevice("R5048", 1);
                                ShowErrMsg("托盘:" + tbTray.Text + textBox5.Text + "存在NC电芯，请检查web日志！");
                            }
                        }

                    }
                    //xieshuju();
                    baocun();
                    tongji();
                    xietongji();
                    isScanned = 0;
                    isWriteType = 0;
                    NGCount = 0;
                    type = -1;
                    count = 0;
                    count2 = 0;
                    adjCount = 0;
                    //isUpload = 0;
                    masterSum = 0;
                    iniFile.WriteValue("XITONG", "needleTimes", needleCount);
                    for (int i = 0; i < 24; i++)
                    {
                        dianya[i] = naizu[i] = wendu[i] = shellVol[i] = 0;
                        jg[i] = "";
                        tb[i].Text = "";
                        tb[i].BackColor = System.Drawing.SystemColors.Control;
                        lb[i].BackColor = System.Drawing.SystemColors.Control;
                    }
                    tbBattay.Text = "";
                    tbPlace.Text = "";
                    tbTray.Text = "";
                    textBox5.Text = "";
                    txbFill.Text = "";
                    comboBox1.Text = "";
                    tbv.Text = "";
                    tbr.Text = "";
                    tbt.Text = "";
                    tbshell.Text = "";
                    lbStatus.Text = "NAN";
                    label22.Text = "当前NG个数：" + NGCount;
                    clearAllData();
                    PutStep++;
                    write();
                    //DelOCV(trayCode);
                }
            }




            //        //cMesForJs mMes = new cMesForJs();
            //        //int iRt =int.Parse(miDataCollectForProcessLotForEach1()); //mMes.TransfOcvData(tb[(ldata[0] - 1)].Text, textBox5.Text, float.Parse(tbv.Text), float.Parse(tbr.Text), float.Parse(textBox3.Text), out sOut);
            //if (plcComm.GetDevice("R0292"))
            //{




            #endregion


            #region  OCV工位/出盘口托盘扫码

            #endregion





            //    //cMesForJs mMes = new cMesForJs();
            //    //string sTesttype = "";
            //    //int iTestway = 0;
            //    //string sErr;
            //    //string[] sCellname = { "1" };//mMes.GetBindInfo(tbTray.Text, out  sTesttype, out iTestway, out sErr);//获取电池条码
            //    //if (sCellname != null)
            //    //{
            //    //    for (int i = 0; i < 24; i++)//显示电池条码到文本框数组
            //    //    {
            //    //        tb[i].Text = sCellname[i];
            //    //    }
            //    //}
            //    //else
            //    //{
            //    //    ShowErrMsg("从系统获取条码异常，正在尝试。。。");
            //    //    return;
            //    //}
            //}
            //catch (Exception err)
            //{
            //    ShowErrMsg(err.Message.ToString());
            //    return;
            //}



            //for (int i = 0; i < tb.Length; i++)
            //{
            //    if (i < 12)
            //    {
            //        strSend = (string.IsNullOrEmpty(tb[i].Text) == true ? "0" : "1") + strSend;//电池一列,1为有电池,0无
            //    }
            //    else
            //    {
            //        strSend1 = (string.IsNullOrEmpty(tb[i].Text) == true ? "0" : "1") + strSend1;//电池二列,1为有电池,0无
            //    }
            //}

        }

        #endregion


        #region Mes调用
        /// <summary>
        /// ocv收数加过站
        /// </summary>
        public dataCollectSfcParametricData miDataCollectForProcessLotForEach1(CollectData[] cd, string cell_code)
        {
            try
            {
                //MiDataCollectForProcessLotForEachServiceImpl cMiDataCollectForProcessLotForEachServiceService = new MiDataCollectForProcessLotForEachServiceImpl().
                //cMiDataCollectForProcessLotForEachServiceService.Timeout = 50000; //超时时间，单位是毫秒。

                //cMiDataCollectForProcessLotForEachServiceService.Url =
                //"http://172.26.11.51:50000/atlmeswebservice/MiDataCollectForProcessLotForEachServiceService?wsdl";
                ////正式测试用地址
                //cMiDataCollectForProcessLotForEachServiceService.Credentials = new System.Net.NetworkCredential(
                //    "44420", "Baiduhai7", null); //验证信息
                //cMiDataCollectForProcessLotForEachServiceService.PreAuthenticate = true;

                dataCollectSfcParametricData cdataCollectSfcParametricData = new dataCollectSfcParametricData();
                cdataCollectSfcParametricData.sfc = cell_code;
                cdataCollectSfcParametricData.dcGroup = "*";
                cdataCollectSfcParametricData.dcGroupRevision = "#";
                ATLocv.MiDataCollectForProcessLotForEach.machineIntegrationParametricData[] cmachineIntegrationParametricDatas = new ATLocv.MiDataCollectForProcessLotForEach.machineIntegrationParametricData[4];
                string para = "电芯：" + cell_code + ";dcGroup:*;dcGroupRevision:#";

                for (int i = 0; i < 3; i++)
                {

                    ATLocv.MiDataCollectForProcessLotForEach.machineIntegrationParametricData cmachineIntegrationParametricData = new ATLocv.MiDataCollectForProcessLotForEach.machineIntegrationParametricData();
                    cmachineIntegrationParametricData.name = cd[i].Attr;
                    cmachineIntegrationParametricData.value = cd[i].Value;
                    cmachineIntegrationParametricData.dataType = ATLocv.MiDataCollectForProcessLotForEach.ParameterDataType.NUMBER;
                    para += "参数：" + cd[i].Attr + "；" + "值：" + cd[i].Value + "；" + "数据类型：" + ATLocv.MiDataCollectForProcessLotForEach.ParameterDataType.NUMBER.ToString();
                    cmachineIntegrationParametricDatas[i] = cmachineIntegrationParametricData;
                }
                ATLocv.MiDataCollectForProcessLotForEach.machineIntegrationParametricData cmachineIntegrationParametricData2 = new ATLocv.MiDataCollectForProcessLotForEach.machineIntegrationParametricData();
                cmachineIntegrationParametricData2.name = cd[3].Attr;
                cmachineIntegrationParametricData2.value = cd[3].Value;
                cmachineIntegrationParametricData2.dataType = ATLocv.MiDataCollectForProcessLotForEach.ParameterDataType.TEXT;
                para += "参数：" + cd[3].Attr + "；" + "值：" + cd[3].Value + "；" + "数据类型：" + ATLocv.MiDataCollectForProcessLotForEach.ParameterDataType.TEXT.ToString();
                saveMESLog(para);
                cmachineIntegrationParametricDatas[3] = cmachineIntegrationParametricData2;

                //ATLocv.MiDataCollectForProcessLotForEach.machineIntegrationParametricData cmachineIntegrationParametricData2 = new ATLocv.MiDataCollectForProcessLotForEach.machineIntegrationParametricData();
                //cmachineIntegrationParametricData2.name = "IMP";
                //cmachineIntegrationParametricData2.value = "3.66";
                //cmachineIntegrationParametricData2.dataType = ATLocv.MiDataCollectForProcessLotForEach.ParameterDataType.NUMBER;

                cdataCollectSfcParametricData.parametricDataArray = cmachineIntegrationParametricDatas;
                return cdataCollectSfcParametricData;
                //return "-1";
            }
            catch (Exception e)
            {
                ShowErrMsg(e.Message.ToString());
                return null;
            }
        }

        public void miDataCollectForProcessLotForEach(dataCollectSfcParametricData[] dcp, string tray_code, int testType, out int code)
        {
            string message = "";
            dataCollectResultArrayData[] result;
            string mesMessage = "";
            MiDataCollectForProcessLotForEachServiceClient mcc = new MiDataCollectForProcessLotForEachServiceClient();


            miDataCollectForProcessLotForEach cmiDataCollectForProcessLotForEach =
                new miDataCollectForProcessLotForEach();

            ATLocv.MiDataCollectForProcessLotForEach.dataCollectForProcessLotForEachRequest cdataCollectForProcessLotForEachRequests = new ATLocv.MiDataCollectForProcessLotForEach.dataCollectForProcessLotForEachRequest();
            cdataCollectForProcessLotForEachRequests.site = "2001";
            cdataCollectForProcessLotForEachRequests.processLot = tray_code;
            if (testType == 1)
            {
                cdataCollectForProcessLotForEachRequests.operation = "TSOCV1";
            }
            if (testType == 3)
            {
                cdataCollectForProcessLotForEachRequests.operation = "TSOCVB";
            }
            cdataCollectForProcessLotForEachRequests.operationRevision = "#";
            cdataCollectForProcessLotForEachRequests.resource = iniFile.GetString("MES", "resource", "");
            cdataCollectForProcessLotForEachRequests.user = "44420";
            cdataCollectForProcessLotForEachRequests.activityID = "?";
            cdataCollectForProcessLotForEachRequests.isDispositionRequired = false;
            cdataCollectForProcessLotForEachRequests.modeProcessSfc = ModeProcessSfc.MODE_PASS_SFC_POST_DC;
            cdataCollectForProcessLotForEachRequests.sfcArray = dcp;

            cmiDataCollectForProcessLotForEach.dataCollectForProcessLotForEachRequest = cdataCollectForProcessLotForEachRequests;
            try
            {
                if (testType == 1)
                {
                    ATLocv.MiDataCollectForProcessLotForEach.dataCollectForProcessLotForEachResponse CmiDataCollectForProcessLotForEachResponse = mcc.miDataCollectForProcessLotForEach("L30_OCV1_1", cdataCollectForProcessLotForEachRequests);
                    message = CmiDataCollectForProcessLotForEachResponse.message;
                    code = CmiDataCollectForProcessLotForEachResponse.code;
                    result = CmiDataCollectForProcessLotForEachResponse.resultArray;
                    
                    mesMessage = string.Format("托盘：" + tbTray.Text + "返回内容：message={0}，code={1},resultArray='{2}'", message, code, result);
                    saveMESLog(mesMessage);
                    saveMESLog("-----------------------------------------------------------------------");
                }
                else
                {
                    ATLocv.MiDataCollectForProcessLotForEach.dataCollectForProcessLotForEachResponse CmiDataCollectForProcessLotForEachResponse = mcc.miDataCollectForProcessLotForEach("L30_OCV1_2", cdataCollectForProcessLotForEachRequests);
                    message = CmiDataCollectForProcessLotForEachResponse.message;
                    code = CmiDataCollectForProcessLotForEachResponse.code;
                    result = CmiDataCollectForProcessLotForEachResponse.resultArray;
                    mesMessage = string.Format("托盘：" + tbTray.Text + "返回内容：message={0}，code={1},resultArray='{2}'", message, code, result);
                    saveMESLog(mesMessage);
                    saveMESLog("-----------------------------------------------------------------------");
                }
            }
            catch (Exception e)
            {
                ShowErrMsg("mes上传数据异常！" + e.Message.ToString());
                code = 1;
            }
            //CmiDataCollectForProcessLotForEachResponse.resultArray.ToString();

        }

        /// <summary>
        /// 获取电芯结果
        /// </summary>
        public string getCellTestResultByTrayId()
        {
            try
            {
                CellTestIntegrationServiceService webSer = new CellTestIntegrationServiceService();
                webSer.Timeout = 50000; //超时时间，单位是毫秒。
                //以下是测试系统webServer地址,正式系统为http://ndmes.catlbattery.com/atlmeswebservice/CellTestIntegrationServiceService?wsdl
                webSer.Url = "http://172.26.11.51:50000/atlmeswebservice/CellTestIntegrationServiceService?wsdl";
                webSer.Credentials = new System.Net.NetworkCredential("44420", "Baiduhai7", null); //验证信息
                webSer.PreAuthenticate = true; //预先身份验证


                getCellTestResultByTrayId gbt = new getCellTestResultByTrayId();
                cellTestResultRequest crr = new cellTestResultRequest();
                crr.site = "2001";
                crr.mode = ModeTrayMatrix.COLUMNFIRST;
                crr.operation = "FORMN3";
                crr.operationRevision = "#";
                crr.processLot = "G3800005";
                crr.resource = "FXXX0001";
                gbt.Request = crr;
                getCellTestResultByTrayIdResponse gCrtr = new getCellTestResultByTrayIdResponse();
                gCrtr = webSer.getCellTestResultByTrayId(gbt);

                gCrtr.@return.code.ToString();
                gCrtr.@return.codeSpecified.ToString();

                gCrtr.@return.message.ToString(); //成功返回null
                sfcResult[] csfcResult = gCrtr.@return.resultArray;
                return gCrtr.@return.code.ToString();

            }
            catch (Exception)
            {
                return null;
            }

        }
        #endregion


        # region 单机SQL语句的增删改查
        private bool UpdateGetPutStatus(string strTableName, int iStatus, int id)//修改盘状态
        {
            string strUpdateCmd = string.Format("update {0} set floor_status='{1}',get_put_time='{2}' where id='{3}'", strTableName, iStatus, DateTime.Now.ToString(), id);
            int iRtn = mSql.ExecuteSql(strUpdateCmd);
            if (iRtn > 0)
                return true;
            else
                return false;
        }

        private bool UpdateGetPutAns(string strTableName, int id)//修改请求
        {
            string strOut = "";
            string strQuerycmd = string.Format("select * from {0} where id='{1}'", strTableName, id);
            DataTable tempTable = mSql.GetDataTableByCmd(strQuerycmd, out strOut);
            int iAns = int.Parse(tempTable.Rows[0]["get_put_req"].ToString());
            DateTime dtimeAns = DateTime.Parse(tempTable.Rows[0]["req_time"].ToString());
            string strUpdateCmd = string.Format("update {0} set get_put_ans='{1}',ans_time='{2}' where id='{3}'", strTableName, iAns, dtimeAns, id);
            int iRtn = mSql.ExecuteSql(strUpdateCmd);
            if (iRtn > 0)
                return true;
            else
                return false;
        }
        #endregion


        #region  统计O1/OB优率
        private void timer1_Tick(object sender, EventArgs e)
        {
            lbtime.Text = DateTime.Now.ToString();
            lbTemp.Text = "温度："+ Convert.ToString((fnormal_temp + fnormal_temp2) / 2);
            for (int i = 0; i < 6; i++)
            {
                sun[i].Text = ((float.Parse(o1[i].Text.ToString())) + (float.Parse(ob[i].Text.ToString()))).ToString();
            }
            float so1 = (float.Parse(o1[0].Text.ToString())) + (float.Parse(o1[1].Text.ToString())) + (float.Parse(o1[2].Text.ToString())) + (float.Parse(o1[3].Text.ToString())) + (float.Parse(o1[4].Text.ToString()));
            float sob = (float.Parse(ob[0].Text.ToString())) + (float.Parse(ob[1].Text.ToString())) + (float.Parse(ob[2].Text.ToString())) + (float.Parse(ob[3].Text.ToString())) + (float.Parse(ob[4].Text.ToString()));
            if (tbo16.Text == "0")
                tbo1yl.Text = "100%";
            else
                tbo1yl.Text = (100 - ((so1 / (float.Parse(tbo16.Text))) * 100)).ToString() + "%";
            if (tbob6.Text == "0")
                tbobyl.Text = "100%";
            else
                tbobyl.Text = (100 - ((sob / (float.Parse(tbob6.Text))) * 100)).ToString() + "%";
            if (tbs6.Text == "0")
                tbsyl.Text = "100%";
            else
                tbsyl.Text = (100 - (((sob + so1) / (float.Parse(tbs6.Text))) * 100)).ToString() + "%";
        }
        #endregion



        #region  统计清零
        private void button5_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("确认统计清零？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1) == DialogResult.Yes)
            {
                lock (obj)
                {
                    tbo11.Text = tbo12.Text = tbo13.Text = tbo14.Text = tbo15.Text = tbo16.Text = "0";
                    tbob1.Text = tbob2.Text = tbob3.Text = tbob4.Text = tbob5.Text = tbob6.Text = "0";
                    textBox1.Text = DateTime.Now.ToString();
                    xietongji();
                }
            }
        }
        #endregion


        #region web端数据交互/啊杰所需数据


        #region 启动与停止
        public bool selectButtonQD()//查找启动按钮的请求   已用
        {
            string strOut = "";
            string strQuerycmd = string.Format("select * from button_action where id='54'");
            DataTable tempTable = mSql.GetDataTableByCmd(strQuerycmd, out strOut);
            int iAns = int.Parse(tempTable.Rows[0]["bt_trigger"].ToString());
            if (iAns == 1)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private bool UpdateButtonQD()//修改启动按钮的请求   已用
        {
            string strUpdateCmd = string.Format("update button_action set bt_trigger='0' where id='54'");
            int iRtn = mSql.ExecuteSql(strUpdateCmd);
            if (iRtn > 0)
                return true;
            else
                return false;
        }

        public bool selectButtonTZ()//查找停止按钮的请求   已用
        {
            string strOut = "";
            string strQuerycmd = string.Format("select * from button_action where id='55'");
            DataTable tempTable = mSql.GetDataTableByCmd(strQuerycmd, out strOut);
            int iAns = int.Parse(tempTable.Rows[0]["bt_trigger"].ToString());
            if (iAns == 1)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private bool UpdateButtonTZ()//修改停止按钮的请求   已用
        {
            string strUpdateCmd = string.Format("update button_action set bt_trigger='0' where id='55'");
            int iRtn = mSql.ExecuteSql(strUpdateCmd);
            if (iRtn > 0)
                return true;
            else
                return false;
        }
        #endregion



        private bool UpdateOCVIsConnectPlc(int IsConnectPlc)//修改PLC连接状态   已用
        {
            string strUpdateCmd = string.Format("update ocvstatus set IsConnectPlc='{0}' where id='1'", IsConnectPlc);
            int iRtn = mysql3.ExecuteSql(strUpdateCmd);
            if (iRtn > 0)
                return true;
            else
                return false;
        }

        private bool AddFacility_run_information(string id, int facility_type_no, int stayguy_no, string information, DateTime createtime)//添加信息提示(facility_run_information表)   已用
        {
            string strUpdateCmd = string.Format("insert into facility_run_information(id,facility_type_no,stayguy_no,information,createtime) values('{0}','{1}','{2}','{3}','{4}')", id, facility_type_no, stayguy_no, information, createtime);
            int iRtn = mysql3.ExecuteSql(strUpdateCmd);
            if (iRtn > 0)
                return true;
            else
                return false;
        }

        private bool AddOCV(string TrayCode, string BatteryCode, int Passagewary, int BatteryState, int id)//添加电池条码
        {
            string strUpdateCmd = string.Format("update ocvdisplay set TrayCode='{0}',BatteryCode='{1}',Passagewary='{2}',BatteryState='{3}' where id='{4}'", TrayCode, BatteryCode, Passagewary, BatteryState, id);
            int iRtn = mysql3.ExecuteSql(strUpdateCmd);
            if (iRtn > 0)
                return true;
            else
                return false;
        }

        private bool DelOCV(string TrayCode)//删除电池条码
        {
            string strUpdateCmd = string.Format("delete from ocvdisplay where TrayCode='{0}'", TrayCode);
            int iRtn = mysql3.ExecuteSql(strUpdateCmd);
            if (iRtn > 0)
                return true;
            else
                return false;
        }

        private bool UpdateOCVTNKC(string batteryCode, float Temperature, float Neizhu, float VoltageValue, int CsPassagewary)//修改内阻，温度，电压，位置
        {
            string strUpdateCmd = string.Format("update ocvstatus set Temperature='{0}',Neizhu='{1}',VoltageValue='{2}',CsPassagewary='{3}',CsCode='{4}' where id='1'", Temperature, Neizhu, VoltageValue, CsPassagewary, batteryCode);
            int iRtn = mysql3.ExecuteSql(strUpdateCmd);
            if (iRtn > 0)
                return true;
            else
                return false;
        }

        private bool UpdateOCVShellVol(float shellVol)//更新壳电压
        {
            string strUpdateSql = string.Format("update ocvstatus set KeVoltage='{0}' where id='1'", shellVol);
            int iRtn = mysql3.ExecuteSql(strUpdateSql);
            if (iRtn > 0)
                return true;
            else
                return false;

        }

        private bool UpdateOCVHaoPorHuaiP(int BatteryState, int Passagewary)//根据通道修改OCV测试结果状态（1良品/2坏品）
        {
            string strUpdateCmd = string.Format("update ocvdisplay set BatteryState='{0}' where id='{1}'", BatteryState, Passagewary);
            int iRtn = mysql3.ExecuteSql(strUpdateCmd);
            if (iRtn > 0)
                return true;
            else
                return false;
        }



        private bool UpdateOCVCanbeputY(int Canbeput)//修改(1有可放信号/0无可放信号)
        {
            string strUpdateCmd = string.Format("update ocvstatus set Canbeput='{0}' where id='1'", Canbeput);
            int iRtn = mysql3.ExecuteSql(strUpdateCmd);
            if (iRtn > 0)
                return true;
            else
                return false;
        }




        private bool UpdateOCVIsTray(int IsTray)//修改(1有盘信号/0无盘信号)
        {
            string strUpdateCmd = string.Format("update ocvstatus set IsTray='{0}' where id='1'", IsTray);
            int iRtn = mysql3.ExecuteSql(strUpdateCmd);
            if (iRtn > 0)
                return true;
            else
                return false;
        }



        private bool UpdateOCVDesirable(int Desirable)//修改(1有可取信号/0无可取信号)
        {
            string strUpdateCmd = string.Format("update ocvstatus set Desirable='{0}' where id='1'", Desirable);
            int iRtn = mysql3.ExecuteSql(strUpdateCmd);
            if (iRtn > 0)
                return true;
            else
                return false;
        }



        private bool UpdateOCVTestCode(string TestCode)//修改测试区托盘条码
        {
            string strUpdateCmd = string.Format("update ocvstatus set TestCode='{0}' where id='1'", TestCode);
            int iRtn = mysql3.ExecuteSql(strUpdateCmd);
            if (iRtn > 0)
                return true;
            else
                return false;
        }


        private bool UpdateOCVMaster(float MasterStandardValues, float MasterErrorRange)//修改Master标准值与误差范围
        {
            string strUpdateCmd = string.Format("update ocvstatus set MasterStandardValues='{0}',MasterErrorRange='{1}' where id='1'", MasterStandardValues, MasterErrorRange);
            int iRtn = mysql3.ExecuteSql(strUpdateCmd);
            if (iRtn > 0)
                return true;
            else
                return false;
        }


        private bool UpdateOCVException(string trayCode, double timeSpan,int handleType)
        {
            string sqlException = string.Format("insert into ocv_handle_exception(tray_code,time_span,error_type) values('{0}','{1}','{2}')", trayCode, timeSpan,handleType);
            int iRtn = mSql.ExecuteSql(sqlException);
            if (iRtn > 0)
                return true;
            else
                return false;
        }
           



        #region  OCV测试数据
        private bool UpdateOCVO1ShuJu(int VoltageBad, int NeiZhu, int Temperature, int Unknown, int Sum, float OptimalRate, int id)//修改O1/OB数据
        {
            string strUpdateCmd = string.Format("update ocvstatistical set VoltageBad='{0}',NeiZhu='{1}',Temperature='{2}',Unknown='{3}',Sum='{4}',OptimalRate='{5}' where id='{6}'", VoltageBad, NeiZhu, Temperature, Unknown, Sum, OptimalRate, id);
            int iRtn = mysql3.ExecuteSql(strUpdateCmd);
            if (iRtn > 0)
                return true;
            else
                return false;
        }


        public bool selectButtonQL()//查找OCV数据清零请求
        {
            string strOut = "";
            string strQuerycmd = string.Format("select * from button_action where id='62'");
            DataTable tempTable = mSql.GetDataTableByCmd(strQuerycmd, out strOut);
            int iAns = int.Parse(tempTable.Rows[0]["bt_trigger"].ToString());
            if (iAns == 1)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private bool UpdateButtonQL()//修改OCV数据清零请求
        {
            string strUpdateCmd = string.Format("update button_action set bt_trigger='0' where id='62'");
            int iRtn = mSql.ExecuteSql(strUpdateCmd);
            if (iRtn > 0)
                return true;
            else
                return false;
        }
        #endregion


        #endregion


        #region  表上传数据/小田所需数据
        private bool AddTestOcv(string battery_code, string tray_id_2, string testType, float temp, float vol, float innerResist, float shellVol, int gallery, DateTime test_time)//添加OCV测试结果到临时表
        {
            string strUpdateCmd = string.Format("insert into TestOcv(battery_code,tray_id_2,testType,temp,vol,innerResist,shellVol,gallery,test_time) values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}')", battery_code, tray_id_2, testType, temp, vol, innerResist, shellVol, gallery, test_time);
            int iRtn = mSql.ExecuteSql(strUpdateCmd);
            if (iRtn > 0)
                return true;
            else
                return false;
        }

        private bool AddOcvData(string battery_code, float temp, float vol, float innerR, float shellVol)
        {
            string addSql = string.Format("update product_statistics set ocv1_temp='{0}',ocv1_vol='{1}',ocv1_innerResist='{2}',ocv1_shellVol='{3}' where battery_code='{4}'", temp, vol, innerR, shellVol, battery_code);
            //string ocvSql = string.Format("insert into product_statistics(ocv1_temp,ocv1_vol,ocv1_innerResist,ocv1_shellVol) values({0},{1},{2},{3}) where battery_code='{4}'", temp, vol, innerR, shellVol, battery_code);
            int iRtn = mysql3.ExecuteSql(addSql);
            if (iRtn > 0)
                return true;
            else
                return false;
        }

        private bool AddOcvDataB(string battery_code, float temp, float vol, float innerR, float shellVol)
        {
            string addSql = string.Format("update product_statistics set ocvb_temp='{0}',ocvb_vol='{1}',ocvb_innerResist='{2}',ocvb_shellVol='{3}' where battery_code='{4}'", temp, vol, innerR, shellVol, battery_code);
            //string ocvSql = string.Format("insert into product_statistics(ocvb_temp,ocvb_vol,ocvb_innerResist,ocvb_shellVol) values({0},{1},{2},{3}) where battery_code='{4}'", temp, vol, innerR, shellVol, battery_code);
            int iRtn = mysql3.ExecuteSql(addSql);
            if (iRtn > 0)
                return true;
            else
                return false;
        }

        private bool AddOcvStatistic(string cell, string tray,float vol, float imp, float temp, DateTime date)
        {
            string sql = string.Format("insert into ocv_statistic(cell_code,tray_code,voltage1,imp1,temperature1,time1) values('{0}','{1}','{2}','{3}','{4}','{5}')", cell,tray,vol, imp, temp, date);
            int iRtn = mSql.ExecuteSql(sql);
            if (iRtn > 0)
                return true;
            else 
                return false;
        }

        private bool AddOcvStatisticB(string cell, float vol, float imp, float temp, DateTime date)
        {
            string sql = string.Format("update ocv_statistic set voltage2='{0}',imp2='{1}',temperature2='{2}',time2='{3}' where cell_code='{4}'", vol, imp, temp, date,cell);
            int iRtn = mSql.ExecuteSql(sql);
            if (iRtn > 0)
                return true;
            else
                return false;
        }

        private bool AddOcvStatisticShell(float shell,string cell)
        {
            string sql = string.Format("update ocv_statistic set shell_vol1='{0}' where cell_code='{1}'", shell, cell);
            int iRtn = mSql.ExecuteSql(sql);
            if (iRtn > 0)
                return true;
            else
                return false;
        }

        private bool AddOcvStatisticShellB(float shell, string cell)
        {
            string sql = string.Format("update ocv_statistic set shell_vol2='{0}' where cell_code='{1}'", shell, cell);
            int iRtn = mSql.ExecuteSql(sql);
            if (iRtn > 0)
                return true;
            else
                return false;
        }

        
        #endregion


        #region  报警值及其产量统计/赵攀所需数据 //现场再确定
        private bool UpdateAlarm_value(int alarm_value, int alarm_content_id)//根据报警ID修改报警触发
        {
            string strUpdateCmd = string.Format("UPDATE alarm_content SET alarm_value='{0}' WHERE alarm_content_id = '{1}'");
            int iRtn = mysql3.ExecuteSql(strUpdateCmd);
            if (iRtn > 0)
                return true;
            else
                return false;
        }

        private bool OperAlarmRecord_proc(int acid, DateTime atime, int av, int facility, int stayguy_no)//调用存储过程
        {
            string strUpdateCmd = string.Format("call operAlarmRecord_proc('{0}','{1}','{2}','{3}','{4}')", acid, DateTime.Now, av, facility, stayguy_no);
            int iRtn = mysql3.ExecuteSql(strUpdateCmd);
            if (iRtn > 0)
                return true;
            else
                return false;
        }

        #endregion


        #region 报警值写入
        private void timer2_Tick(object sender, EventArgs e)
        {

            if (webtrue)
            {
                string strOut = "";
                string strQuerycmd = string.Format("select * from button_action");
                DataTable tempTable = mysql3.GetDataTableByCmd(strQuerycmd, out strOut);
                if (tempTable != null)
                {
                    webtrue = true;
                }
                else
                {
                    webtrue = false;
                    ShowErrMsg("连接WEB端数据库异常，请检查网络端口等是否正常！");
                }
            }


            if (webtrue)
            {
                if (selectButtonQD())//查找启动按钮，若为1，则启动
                {
                    UpdateButtonQD();//将bt_trigger字段设为0
                    if (button2.Text == "运行")
                    {
                        try
                        {
                            string sOut = "";
                            DataTable tempTable = mSql.GetDataTableByCmd("select * from sys_floor_table2", out sOut);
                            if (tempTable == null)
                            {
                                ShowErrMsg("数据库连接失败，" + sOut);
                                return;
                            }
                            if (!sp3562.IsOpen)//内阻表串口连接

                                sp3562.Open();

                            //cgVar.mLoad.portOpen();//托盘扫码打开串口

                            gettemperature_normal();//温度表串口连接

                            sp3562.ReadExisting();
                            sp3562.WriteLine("*IDN?");

                            Thread.Sleep(100);
                            string bm = sp3562.ReadExisting();
                            if (bm == "")
                            {
                                ShowErrMsg("与3562表通讯异常");
                                return;
                            }

                            sp3562.WriteLine(":RESISTANCE:RANGE 3.0000E-3");
                            if (ctTest1 == null)//电压表网口连接
                            {
                                ctTest1 = new clientTest(strIP, iPort);
                                ctTest1.connectServer();
                            }
                            if (!ctTest1.isConnected)
                            {
                                ShowErrMsg("与34461表通讯异常");
                                return;
                            }
                            if (fnormal_temp == 0)
                            {
                                ShowErrMsg("与测温仪通讯异常");
                                return;
                            }


                            if (string.IsNullOrEmpty(strIP2) || iPort2 == 0)
                            {
                                ShowErrMsg("PLC通信失败，请检查网络通信地址端口等是否正常！");
                                return;
                            }

                            if (plcNetOperator == null)
                                plcNetOperator = new clientTest(strIP2, iPort2);//"192.168.10.202", 9094);

                            if (!plcNetOperator.isConnected)
                                plcNetOperator.connectServer();

                            Thread.Sleep(1000);

                            if (!plcNetOperator.isConnected)
                            {
                                ShowErrMsg("PLC通信失败，请检查网络接口时候正常！");
                                return;
                            }

                            plcComm.plc = plcNetOperator;

                            // if (plcoperator.setdevice("r0002", 1) == 0)//ocv测试机启动信号
                            //{
                            if (plcComm.SetDevice("R0332", 1) > 0)
                            {
                                if (thOCVLive == null)
                                {
                                    thOCVLive = new Thread(new ThreadStart(this.RpjLive));
                                    thOCVLive.IsBackground = true;
                                    thOCVLive.Start();
                                }

                                timerun.Enabled = true;
                                button2.Text = "停止";
                                label18.BackColor = Color.Green;
                                UpdateOCVIsConnectPlc(1);
                            }
                            //return;
                            // }
                            //else
                            //{
                            //    ShowErrMsg("与PLC通讯失败！");
                            //    return;
                            //}
                        }
                        catch (Exception err)
                        {
                            ShowErrMsg(err.Message.ToString());
                        }
                    }
                    else
                    {
                        ShowErrMsg("设备处在停止状态！");
                    }
                }
                //else
                //{

                //    //if (plcOperator.SetDevice("R0002", 0) == 0)
                //    //{
                //    if (thOCVLive != null)
                //    {
                //        thOCVLive.Abort();
                //        thOCVLive = null;
                //    }
                //    button2.Text = "运行";
                //    timerun.Enabled = false; ;
                //    //cgVar.mLoad.portClose();//串口关闭
                //    UpdateOCVIsConnectPlc(1);
                //    //}
                //}
                if (selectButtonTZ())
                {
                    UpdateButtonTZ();
                    if (thOCVLive != null)
                    {
                        thOCVLive.Abort();
                        thOCVLive = null;
                    }
                    button2.Text = "运行";
                    plcComm.SetDevice("R0332", 0);
                    timerun.Enabled = false;

                    if (plcNetOperator != null)
                    {
                        plcNetOperator = null;
                    }
                    label18.BackColor = Color.Gray;
                    UpdateOCVIsConnectPlc(0);
                }
                if (selectButtonQL())
                {
                    UpdateButtonQL();
                    for (int i = 0; i < 6; i++)
                    {
                        o1[i].Text = "0";
                        ob[i].Text = "0";
                        sun[i].Text = "0";

                    }
                    tbo1yl.Text = "100%";
                    tbobyl.Text = "100%";
                    tbsyl.Text = "100%";
                }

            }




            //int[] ldata = { 0, 0 };
            //if (plcOperator.ReadDeviceBlock("901", 1, out ldata) == 0)
            //{
            //    string alm = Convert.ToString(ldata[0], 2);//十进制转二进制
            //    for (int i = 0; i < alm.Length; i++)
            //    {
            //        string al = alm.Substring(i, 1);//截取报警位置的字符串
            //        if (al == "1")
            //        {
            //            int weizhi = i + 1;
            //            //写入对应报警位
            //            UpdateAlarm_value(1, weizhi);

            //            if (pList.ContainsKey(weizhi.ToString()))//判断该键值是否已经存在
            //            {
            //                ShowErrMsg("该报警信息已存在");
            //            }
            //            else
            //            {
            //                //调用报警存储过程
            //                if (OperAlarmRecord_proc(weizhi, DateTime.Now, 1, 4, 1))
            //                {
            //                    pList.Add(weizhi.ToString(), "1");
            //                }
            //            }
            //        }
            //        else
            //        {
            //            int weizhi = i + 1;
            //            //对应位置不报警
            //            UpdateAlarm_value(0, weizhi);
            //            if (pList.ContainsKey(weizhi.ToString()))
            //            {
            //                //调用报警存储过程
            //                if (OperAlarmRecord_proc(weizhi, DateTime.Now, 0, 4, 1))
            //                {
            //                    pList.Remove(weizhi.ToString());
            //                }
            //            }
            //        }
            //    }
            //}
        }


        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            //int[] ldata ={ 2 };
            //string[] dd = null;
            //string ada = "";
            // plcComm.WriteDeviceBlock("32500", 1, ldata);
            // plcComm.ReadMoreDevice2(out ada, dd);

            //int code = 0;
            //dataCollectSfcParametricData[] dcp = new dataCollectSfcParametricData[24];
            //for (int n = 0; n < 24; n++)
            //{
            //    CollectData[] cCollectData = new CollectData[] {new CollectData("OCV",dianya[n].ToString()),new CollectData("IMP",naizu[n].ToString()),
            //                        new CollectData("TEMPERATURE",wendu[n].ToString()),new CollectData("CHANNELID",(n+1).ToString())};
            //    //dataList.Add(miDataCollectForProcessLotForEach1(cCollectData, tbTray.Text));
            //    dcp[n] = miDataCollectForProcessLotForEach1(cCollectData, "11111");
            //}
            //miDataCollectForProcessLotForEach(dcp, tbTray.Text, type, out code);
            //if (code == 0)
            //    ShowErrMsg("托盘:" + tbTray.Text + "电芯数据上传MES成功！");
            //else
            //{
            //    timerun.Enabled = false;
            //    if (MessageBox.Show("托盘:" + tbTray.Text + "电芯数据上传MES失败！", "警告", MessageBoxButtons.RetryCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.Cancel)
            //    {
            //        //发信号
            //        timerun.Enabled = true;
            //    }
            //    else
            //    {
            //        timerun.Enabled = true;
            //        int i = 0;
            //        while (i < 3)
            //        {
            //            miDataCollectForProcessLotForEach(dcp, tbTray.Text, type, out code);
            //            if (code == 0)
            //            {
            //                break;
            //            }
            //            i++;
            //        }
            //        if (i <= 1)
            //        {
            //            ShowErrMsg("托盘:" + tbTray.Text + "电芯数据上传MES成功！");
            //        }
            //        else
            //            ShowErrMsg("托盘:" + tbTray.Text + "电芯数据上传MES失败！");
            //    }
            //}
            plcComm.SetDevice("R5048", 1);
        }


        #region 手动清零
        private void btnClear_Click(object sender, EventArgs e)
        {
            if (plcComm.SetDevice("R0270", 1) > 0)
            {
                ShowErrMsg("准备清零……");
                Thread.Sleep(1000);
                if (plcComm.GetDevice("R0271"))//仪表清零到位信号
                {
                    ShowErrMsg("仪表清零开始");
                    // Thread.Sleep(10000);
                    sp3562.WriteLine(":ADJ?");
                    Thread.Sleep(1000);
                    string sRtn = sp3562.ReadExisting();
                    tbr.Text = sRtn;
                    if (sRtn.Trim() == "0")
                    {
                        plcComm.SetDevice("R0272", 1);//仪表清零校准成功
                        ShowErrMsg("清零完成");
                        button3.Enabled = true;
                    }
                }
                else
                {
                    // return;
                    ShowErrMsg("发送清零信号失败");
                }
            }
        }
        #endregion


        #region 手动内阻校验
        private void button3_Click(object sender, EventArgs e)
        {

        }
        #endregion

        #region 重设清零次数
        private void button6_Click(object sender, EventArgs e)
        {

        }
        #endregion


    }


}
