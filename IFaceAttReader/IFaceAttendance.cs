using System;

namespace IFaceAttReader
{
    public class IFaceAttendance
    {
        public string EnrollNumber { get; set; }
        public int IsInValid { get; set; }
        public int AttState { get; set; }
        public int VerifyMethod { get; set; }
        public int WorkCode { get; set; }
        public DateTime Time { get; set; }
        public string deviceName { get; set; }

        public IFaceAttendance(){}

        public IFaceAttendance(string EnrollNumber, int IsInValid, int AttState, int VerifyMethod, int WorkCode, DateTime Time, string deviceName)
        {
            this.EnrollNumber = EnrollNumber;
            this.AttState = AttState;
            this.IsInValid = IsInValid;
            this.VerifyMethod = VerifyMethod;
            this.WorkCode = WorkCode;
            this.Time = Time;
            this.deviceName = deviceName;
        }
    }
}
