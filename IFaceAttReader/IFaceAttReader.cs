using System;
using System.Collections.Generic;
using System.Configuration;
using System.ServiceProcess;
using System.Threading;
using Newtonsoft.Json;
using System.Windows.Forms;

namespace IFaceAttReader
{
    public partial class IFaceAttReader : ServiceBase
    {
        //the serial number of the device.After connecting the device ,this value will be changed.
        private int iMachineNumber = 1;

        public static string IFaceDevices = ConfigurationManager.AppSettings.Get("IFaceDevices");
        public static string IFaceCheckTime = ConfigurationManager.AppSettings.Get("IFaceCheckTime");
        public static string IFaceCheckInterval = ConfigurationManager.AppSettings.Get("IFaceCheckInterval");

        Dictionary<string, HashSet<string>> dictionary = new Dictionary<string, HashSet<string>>();
        List<string[]> checkTimes = new List<string[]>();
        int idwErrorCode = 0;

        public IFaceAttReader()
        {
            InitializeComponent();

            // from https://www.codeproject.com/Questions/711973/Using-Zkemkeeper-dll-from-SDK-for-Biometric-scanne
            string[] devices = IFaceDevices.Split(';');

            if (devices.Length <= 0)
            {
                LogHelper.Log(LogLevel.Debug, "device IP & port param set error!");
                return;
            }

            // handle times
            string[] check_times = IFaceCheckTime.Split(';');
            if (check_times.Length <= 0)
            {
                LogHelper.Log(LogLevel.Debug, "check times param set error!");
                return;
            }

            foreach (string time in check_times)
            {
                string[] sub_time = time.Split('-');
                if (sub_time.Length != 2)
                {
                    LogHelper.Log(LogLevel.Debug, "check time param set error: " + time);
                    continue;
                }
                string[] time_pair = {sub_time[0] ,sub_time[1]};
                checkTimes.Add(time_pair);
            }

            
            foreach (string device in devices)
            {
                string[] ip_port_passwd = device.Split('@');
                if (ip_port_passwd.Length != 2)
                {
                    LogHelper.Log(LogLevel.Debug, "device IP_port & commKey param set error! :" + device);
                    continue;
                }
                string ip_port_str = ip_port_passwd[0];
                string commKey = ip_port_passwd[1];
                string[] ip_port = ip_port_str.Split(':');
                if (ip_port.Length != 2)
                {
                    LogHelper.Log(LogLevel.Debug, "device IP & port param set error! :" + ip_port_str);
                    continue;
                }
                ConnectDevice(ip_port[0], int.Parse(ip_port[1]), int.Parse(commKey));
                
            }

            queryAttRecords_3_days();
        }

        public void queryAttRecords_3_days()
        {
            // get 3 day's reords
            DateTime time_3_days_ago = DateTime.Now.AddDays(-3);
            String time_3_days_ago_str = time_3_days_ago.ToString("yyyy-MM-dd");

            List<IFaceAttendance> records = gueryLatestAttData(time_3_days_ago_str);
            foreach (IFaceAttendance att in records)
            {
                if (!dictionary.ContainsKey(att.deviceName))
                {
                    HashSet<string> newSet = new HashSet<string>();
                    dictionary.Add(att.deviceName, newSet);
                }
                HashSet<string> set = dictionary[att.deviceName];
                string record = att.EnrollNumber + "@" + att.Time.ToString("yyyy-MM-dd HH:mm:ss");
                set.Add(record);
                LogHelper.Log(LogLevel.Debug, "record in database: " + record + "by " + att.deviceName);
            }
            LogHelper.Log(LogLevel.Debug, "records in database 3days total : " +  records.Count);
        }

        public void ConnectDevice(string iface_Ip, int port, int commKey) {
            string deviceName = iface_Ip + "_" + port;
            if (! dictionary.ContainsKey(deviceName))
            {
                HashSet<string> set = new HashSet<string>();
                dictionary.Add(deviceName, set);
            }
            Thread createComAndMessagePumpThread = new Thread(() =>
            {
                int machineNumber = 1;
                LogHelper.Log(LogLevel.Debug, "connectting to device:" + deviceName);
                zkemkeeper.CZKEMClass axCZKEM1 = new zkemkeeper.CZKEMClass();
                if (commKey != 0)
                {
                    axCZKEM1.SetCommPassword(commKey);
                }
                System.Timers.Timer timer = new System.Timers.Timer();
                timer.Interval = int.Parse(IFaceCheckInterval);
                timer.Enabled = true;
                connect(axCZKEM1, iface_Ip, port, 0);
                timer.Elapsed += new System.Timers.ElapsedEventHandler((object sender, System.Timers.ElapsedEventArgs e) => 
                {
                    DateTime now = DateTime.Now;
                    // compare time
                    foreach (string[] time_pair in checkTimes)
                    {
                        if (now >= Convert.ToDateTime(time_pair[0]) && now <= Convert.ToDateTime(time_pair[1]))
                        {
                            string IPAddr = "";
                            int retry_times = 1;
                            while (! axCZKEM1.GetDeviceIP(machineNumber, IPAddr))
                            {
                                axCZKEM1.GetLastError(ref idwErrorCode);
                                if (retry_times > 2) // only retry 2 times
                                {
                                    retry_times--;
                                    LogHelper.Log(LogLevel.Debug, "Connecting to " + deviceName + " failed " + retry_times + " times, ErrorCode=" + idwErrorCode.ToString() + ", stop connecting.");
                                    return;
                                }
                                LogHelper.Log(LogLevel.Debug, "Connecting to " + deviceName + " failed, ErrorCode=" + idwErrorCode.ToString() + ", reConnecting: " + retry_times);
                                connect(axCZKEM1, iface_Ip, port, retry_times);
                                retry_times++;
                            }
                            LogHelper.Log(LogLevel.Debug, "device " + deviceName + " connected status is ok.");
                            readAttData(axCZKEM1, iface_Ip + "_" + port);
                        }
                        else{
                            
                            LogHelper.Log(LogLevel.Debug, "device " + deviceName + " time not in " + time_pair[0] + "--" + time_pair[1] + ", continued.");
                            continue;
                        }
                    }
                });
                Application.Run();
            });
            createComAndMessagePumpThread.SetApartmentState(ApartmentState.STA);
            createComAndMessagePumpThread.Name = deviceName;
            createComAndMessagePumpThread.IsBackground = true;
            createComAndMessagePumpThread.Start();

            LogHelper.Log(LogLevel.Debug , deviceName + ": Thread Started.");
        }


        private void connect(zkemkeeper.CZKEMClass axCZKEM1,string iface_Ip, int port, int retry)
        {
            bool bIsConnected = axCZKEM1.Connect_Net(iface_Ip, port);
            string deviceName = Thread.CurrentThread.Name;
            if (bIsConnected == true){
                LogHelper.Log(LogLevel.Debug, "Connected " + deviceName + " successed by retrying " + retry + " times.");
                iMachineNumber = 1;//In fact,when you are using the tcp/ip communication,this parameter will be ignored,that is any integer will all right.Here we use 1.
                bool regEvent = axCZKEM1.RegEvent(iMachineNumber, 65535);//Here you can register the realtime events that you want to be triggered(the parameters 65535 means registering all)
                if (regEvent == true)
                {
                    LogHelper.Log(LogLevel.Debug, deviceName +  " regEvent value: " + regEvent);
                    axCZKEM1.OnAttTransactionEx += new zkemkeeper._IZKEMEvents_OnAttTransactionExEventHandler(axCZKEM1_OnAttTransactionEx);
                }
                else
                {
                    LogHelper.Log(LogLevel.Debug, deviceName + " regEvent failed, disconnecting device...");
                    axCZKEM1.Disconnect();
                }
            }
            else
            {
                axCZKEM1.GetLastError(ref idwErrorCode);
                axCZKEM1.OnAttTransactionEx -= new zkemkeeper._IZKEMEvents_OnAttTransactionExEventHandler(axCZKEM1_OnAttTransactionEx);
                LogHelper.Log(LogLevel.Debug, "Unable to connect " + iface_Ip + "_" + port + " by retrying " + retry + " times. ErrorCode=" + idwErrorCode.ToString() + ", connect failed.");
            }

        }

        private void axCZKEM1_OnAttTransactionEx(string sEnrollNumber, int iIsInValid, int iAttState, int iVerifyMethod, int iYear, int iMonth, int iDay, int iHour, int iMinute, int iSecond, int iWorkCode)
        { 
            string time = iYear.ToString() + "-" 
                        + (iMonth < 10 ? "0" : "") + iMonth.ToString() + "-" 
                        + (iDay < 10 ? "0" : "") + iDay.ToString() + " "
                        + (iHour < 10 ? "0" : "") + iHour.ToString() + ":"
                        + (iMinute < 10 ? "0" : "") + iMinute.ToString() + ":"
                        + (iSecond < 10 ? "0" : "") + iSecond.ToString();
            string deviceName = Thread.CurrentThread.Name;
            HashSet<string> set = dictionary[deviceName];
            LogHelper.Log(LogLevel.Debug, "Teacher " + sEnrollNumber + " attendance @" + time + " by " + deviceName + " event.");
            DateTime recordTime = Convert.ToDateTime(time);
            int dataSize = SaveAttData(new IFaceAttendance(sEnrollNumber, iIsInValid, iAttState , iVerifyMethod , iWorkCode ,recordTime, deviceName));
            if(dataSize > 0){
                set.Add(sEnrollNumber + "@" + time);
                LogHelper.Log(LogLevel.Debug, sEnrollNumber + "@" + time + " by " + deviceName + "add to set, save size: " + dataSize);
            }else{
                LogHelper.Log(LogLevel.Debug, sEnrollNumber + "@" + time + " by " + deviceName + "not add to set, save size: " + dataSize);
            }
        }

        private void readAttData(zkemkeeper.CZKEMClass axCZKEM1, string deviceName)
        {
            //axCZKEM1.EnableDevice(iMachineNumber, false);//disable the device
            LogHelper.Log(LogLevel.Debug, "Check attendence records for " + deviceName);
            int count = 0;
            int repeat = 0;
            int total = 0;
            if (axCZKEM1.ReadGeneralLogData(iMachineNumber))//read all the attendance records to the memory
            {
                string sdwEnrollNumber = "";
                int idwVerifyMode = 0;
                int idwInOutMode = 0;
                int idwYear = 0;
                int idwMonth = 0;
                int idwDay = 0;
                int idwHour = 0;
                int idwMinute = 0;
                int idwSecond = 0;
                int idwWorkcode = 0;
                HashSet<string> set = dictionary[deviceName];
                string time_3days_ago_str = DateTime.Now.AddDays(-3).ToString("yyyy-MM-dd") + " 00:00:00";
                DateTime time_3days_ago = Convert.ToDateTime(time_3days_ago_str);
                while (axCZKEM1.SSR_GetGeneralLogData(iMachineNumber, out sdwEnrollNumber, out idwVerifyMode,
                            out idwInOutMode, out idwYear, out idwMonth, out idwDay, out idwHour, out idwMinute, out idwSecond, ref idwWorkcode))//get records from the memory
                {
                    total++;
                    string time = idwYear.ToString() + "-" 
                        + (idwMonth < 10 ? "0" : "") + idwMonth.ToString() + "-" 
                        + (idwDay < 10 ? "0" : "") + idwDay.ToString() + " "
                        + (idwHour < 10 ? "0" : "") + idwHour.ToString() + ":"
                        + (idwMinute < 10 ? "0" : "") + idwMinute.ToString() + ":"
                        + (idwSecond < 10 ? "0" : "") + idwSecond.ToString();
                    DateTime recordTime = Convert.ToDateTime(time);
                    string record = sdwEnrollNumber + "@" + time;
                    if (recordTime < time_3days_ago)
                    {
                        if(set.Contains(record)){
                            set.Remove(record);
                        }
                        continue;
                    }
                    if (set.Contains(record))
                    {
                        repeat++;
                        continue;
                    }
                    else
                    {
                        set.Add(record);
                        count++;
                        LogHelper.Log(LogLevel.Debug, "Teacher " + sdwEnrollNumber + " attendance @" + time + " by " + deviceName + " check.");
                        SaveAttData(new IFaceAttendance(sdwEnrollNumber, 0, idwInOutMode, idwVerifyMode, idwWorkcode, recordTime, deviceName));
                    }  
                }
            }
            LogHelper.Log(LogLevel.Debug, count + " records checked, " + repeat + " repeat, total " + total + " for " + deviceName);
            //axCZKEM1.EnableDevice(iMachineNumber, true);//enable the device
        }

        public int SaveAttData(IFaceAttendance attData)
        {
            #region   插入单条数据
            string sql = @"insert into iface_attendance_record  
            (EnrollNumber, IsInValid, AttState, VerifyMethod, WorkCode, Time, deviceName) values 
            (@EnrollNumber, @IsInValid, @AttState, @VerifyMethod, @WorkCode, @Time, @deviceName)";
            try
            {

            var result = DapperDBContext.Execute(sql, attData);
            if (result >= 1)
            {
                LogHelper.Log(LogLevel.Debug, "Save att success:" + JsonConvert.SerializeObject(attData));
            }
            return result;
            }
            catch (Exception e)
            {
                LogHelper.Log(LogLevel.Fatal, "Save att failed:" + JsonConvert.SerializeObject(attData) + "with error:" +  e.Message);
                return 0;
            }
            #endregion
        }

        public List<IFaceAttendance> gueryLatestAttData(string time)
        {
            #region 
            string sql = @"select * from iface_attendance_record where Time > @Time";
            try
            {

                var result = DapperDBContext.Query(sql, time);
                if (result.Count > 0)
                {
                    LogHelper.Log(LogLevel.Debug, "Get today's att success: " + result.Count);
                }
                return result;
            }
            catch (Exception e)
            {
                LogHelper.Log(LogLevel.Fatal, "Get att failed:" + "with error:" +  e.Message);
                return new List<IFaceAttendance>();
            }
            #endregion
        }


        protected override void OnStart(string[] args)
        {
            LogHelper.Log(LogLevel.Debug, "service started. ");
        }

        protected override void OnStop()
        {
            LogHelper.Log(LogLevel.Debug, "service stopping...");
            
        }
    }
}
