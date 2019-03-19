using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Windows.Forms;

namespace IFaceAttReader
{
    public partial class IFaceAttReader : ServiceBase
    {
        //the serial number of the device.After connecting the device ,this value will be changed.
        private int iMachineNumber = 1;

        public static string IFaceDevices = ConfigurationManager.AppSettings.Get("IFaceDevices");

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
                    LogHelper.Log(LogLevel.Debug, "device IP & port param set error!");
                    continue;
                }
                ConnectDevice(ip_port[0], int.Parse(ip_port[1]), int.Parse(commKey));
            }
        }


        public void ConnectDevice(string iface_Ip, int port, int commKey) { 
            Thread createComAndMessagePumpThread = new Thread(() =>
            {
                LogHelper.Log(LogLevel.Debug, "connectting to device:" + iface_Ip + ":" + port);
                zkemkeeper.CZKEMClass axCZKEM1 = new zkemkeeper.CZKEMClass();
                axCZKEM1.SetCommPassword(commKey);
                ReConnect(axCZKEM1, iface_Ip, port);

                System.Timers.Timer timer = new System.Timers.Timer();
                timer.Elapsed += new System.Timers.ElapsedEventHandler((object sender, System.Timers.ElapsedEventArgs e) => 
                {
                    string IPAddr = "";
                    if (axCZKEM1.GetDeviceIP(iMachineNumber, IPAddr))
                    {
                        LogHelper.Log(LogLevel.Debug, "device " + iface_Ip + ":" + port + " connect status is ok.");
                    }
                    else
                    {
                        axCZKEM1.GetLastError(ref idwErrorCode);
                        axCZKEM1.OnAttTransactionEx -= new zkemkeeper._IZKEMEvents_OnAttTransactionExEventHandler(axCZKEM1_OnAttTransactionEx);
                        LogHelper.Log(LogLevel.Debug, "Unable to connect the device,ErrorCode=" + idwErrorCode.ToString() + ", reConnecting...");
                        ReConnect(axCZKEM1, iface_Ip, port);

                    }
                });
                timer.Interval = 600000;// 10 minutes
                timer.Enabled = true;

                Application.Run();
            });
            createComAndMessagePumpThread.SetApartmentState(ApartmentState.STA);
            createComAndMessagePumpThread.Name = iface_Ip+"_"+port;
            createComAndMessagePumpThread.IsBackground = true;
            createComAndMessagePumpThread.Start();

            LogHelper.Log(LogLevel.Debug , iface_Ip + "_" + port + ": Thread Started.");
        }


        private void ReConnect(zkemkeeper.CZKEMClass axCZKEM1,string iface_Ip, int port)
        {
            bool bIsConnected = axCZKEM1.Connect_Net(iface_Ip, port);

            if (bIsConnected == true){
                iMachineNumber = 1;//In fact,when you are using the tcp/ip communication,this parameter will be ignored,that is any integer will all right.Here we use 1.
                bool regEvent = axCZKEM1.RegEvent(iMachineNumber, 65535);//Here you can register the realtime events that you want to be triggered(the parameters 65535 means registering all)
                LogHelper.Log(LogLevel.Debug, "Connect the device successed. iMachineNumber: " + iMachineNumber);
                if (regEvent == true)
                {
                    LogHelper.Log(LogLevel.Debug, "regEvent value: " + regEvent);
                    axCZKEM1.OnAttTransactionEx += new zkemkeeper._IZKEMEvents_OnAttTransactionExEventHandler(axCZKEM1_OnAttTransactionEx);
                }
                else
                {
                    LogHelper.Log(LogLevel.Debug, "regEvent failed, disconnecting device...");
                    axCZKEM1.Disconnect();
                }
            }
            else
            {
                axCZKEM1.GetLastError(ref idwErrorCode);
                axCZKEM1.OnAttTransactionEx -= new zkemkeeper._IZKEMEvents_OnAttTransactionExEventHandler(axCZKEM1_OnAttTransactionEx);
                LogHelper.Log(LogLevel.Debug, "Unable to connect the device,ErrorCode=" + idwErrorCode.ToString() + ", reConnect failed.");
            }

        }


        private void axCZKEM1_OnAttTransactionEx(string sEnrollNumber, int iIsInValid, int iAttState, int iVerifyMethod, int iYear, int iMonth, int iDay, int iHour, int iMinute, int iSecond, int iWorkCode)
        { 
            string time = iYear.ToString() + "-" + iMonth.ToString() + "-" + iDay.ToString() + " " + iHour.ToString() + ":" + iMinute.ToString() + ":" + iSecond.ToString();
            string deviceName = Thread.CurrentThread.Name;
            LogHelper.Log(LogLevel.Debug, "Teacher " + sEnrollNumber + " attendance @" + time + " by " + deviceName);
            int dataSize = SaveAttData(new IFaceAttendance(sEnrollNumber, iIsInValid, iAttState , iVerifyMethod , iWorkCode ,time, deviceName));
        }


        public int SaveAttData(IFaceAttendance attData)
        {
            #region   插入单条数据
            string sql = @"insert into iface_attendance_record  
            (EnrollNumber, IsInValid, AttState, VerifyMethod, WorkCode, Time, deviceName) values 
            (@EnrollNumber, @IsInValid, @AttState, @VerifyMethod, @WorkCode, @Time, @deviceName)";
            var result = DapperDBContext.Execute(sql, attData); //直接传送list对象
            if (result >= 1)
            {
                LogHelper.Log(LogLevel.Debug, ": save att :" + JsonConvert.SerializeObject(attData));
            }
            return result;
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
