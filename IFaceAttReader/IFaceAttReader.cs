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
       
        //the boolean value identifies whether the device is connected
        private bool bIsConnected = false;
        //the serial number of the device.After connecting the device ,this value will be changed.
        private int iMachineNumber = 1;

        public static readonly string port = ConfigurationManager.AppSettings.Get("IFace_port");
        //IFace_IP
        public static readonly string iface_Ip = ConfigurationManager.AppSettings.Get("IFace_IP");
        //Create Standalone SDK class dynamicly.
        public zkemkeeper.CZKEMClass axCZKEM1 = new zkemkeeper.CZKEMClass();


        public IFaceAttReader()
        {
            InitializeComponent();
            LogHelper.Log(LogLevel.Debug, "connectting to device:" + iface_Ip + ":" + port);
            
            // from https://www.codeproject.com/Questions/711973/Using-Zkemkeeper-dll-from-SDK-for-Biometric-scanne
            Thread createComAndMessagePumpThread = new Thread(() =>
            {
                bIsConnected = axCZKEM1.Connect_Net(iface_Ip, Int16.Parse(port));

                ConnectDevice();

                Application.Run();
            });
            createComAndMessagePumpThread.SetApartmentState(ApartmentState.STA);

            createComAndMessagePumpThread.Start();

            LogHelper.Log(LogLevel.Debug ,"Service Started");
        }



        private void ConnectDevice()
        {
            int idwErrorCode = 0;
            
            if (bIsConnected == true)
            {
                iMachineNumber = 1;//In fact,when you are using the tcp/ip communication,this parameter will be ignored,that is any integer will all right.Here we use 1.
                bool regEvent = axCZKEM1.RegEvent(iMachineNumber, 65535);//Here you can register the realtime events that you want to be triggered(the parameters 65535 means registering all)
                LogHelper.Log(LogLevel.Debug, "Connect the device successed: iMachineNumber: " + iMachineNumber);
                if (regEvent == true)
                {
                    LogHelper.Log(LogLevel.Debug, "regEvent value : " + regEvent);
                    this.axCZKEM1.OnAttTransactionEx += new zkemkeeper._IZKEMEvents_OnAttTransactionExEventHandler(axCZKEM1_OnAttTransactionEx);

                }
                else
                {
                    LogHelper.Log(LogLevel.Debug, "regEvent failed, stop app and disconnecting device...");
                    axCZKEM1.Disconnect();
                }

            }
            else
            {
                axCZKEM1.GetLastError(ref idwErrorCode);
                this.axCZKEM1.OnAttTransactionEx -= new zkemkeeper._IZKEMEvents_OnAttTransactionExEventHandler(axCZKEM1_OnAttTransactionEx);
                LogHelper.Log(LogLevel.Debug, "Unable to connect the device,ErrorCode=" + idwErrorCode.ToString(), "Error");
                return;
            }
        }
        
        private void axCZKEM1_OnAttTransactionEx(string sEnrollNumber, int iIsInValid, int iAttState, int iVerifyMethod, int iYear, int iMonth, int iDay, int iHour, int iMinute, int iSecond, int iWorkCode)
        {
            string time = iYear.ToString() + "-" + iMonth.ToString() + "-" + iDay.ToString() + " " + iHour.ToString() + ":" + iMinute.ToString() + ":" + iSecond.ToString();
            int dataSize = SaveAttData(new IFaceAttendance(sEnrollNumber, iIsInValid, iAttState , iVerifyMethod , iWorkCode ,time));
        }

        


        public int SaveAttData(IFaceAttendance attData)
        {
            #region   插入单条数据
            string sql = @"insert into test.attendance_data 
            (EnrollNumber, IsInValid, AttState, VerifyMethod, WorkCode, Time) values 
            (@EnrollNumber, @IsInValid, @AttState, @VerifyMethod, @WorkCode, @Time)";
            var result = DapperDBContext.Execute(sql, attData); //直接传送list对象
            if (result >= 1)
            {
                LogHelper.Log(LogLevel.Debug, "save att :" + JsonConvert.SerializeObject(attData));

            }
            return result;
            #endregion
        }



        protected override void OnStart(string[] args)
        {
            LogHelper.Log(LogLevel.Debug, "started. bIsConnected is: " + bIsConnected);
        }

        protected override void OnStop()
        {
            LogHelper.Log(LogLevel.Debug, "stopping...: " + bIsConnected);
            axCZKEM1.Disconnect();
        }
    }
}
