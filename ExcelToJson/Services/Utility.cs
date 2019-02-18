using System;
using System.Collections;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Threading;
using System.Web;
using System.Text;

/// <summary>
/// Summary description for Utility
/// </summary>
namespace ExcelToJson.Controllers
{
    public class Utility
    {
        public static string appIp = "localhost/";// System.Net.Dns.Resolve(System.Net.Dns.GetHostName()).AddressList[0].ToString();
        
        public static string appHostCode = "APP_ROOT_URL";
        private enum enmBtnText { Save, Update, Cancel, Clear, Add }


        
        //private readonly IUserInfo iUserInfo;

        public Utility()
        {
            
        }

        //public bool CheckIsNofificationEnable(short eventId, long userId, string notificationType)//notification type : C=Computer, E=EMAIL, S=SMS
        //{
        //    bool result = iNtf.CheckIsNofificationEnable(eventId, userId, notificationType);
        //    return result;
        //}

        //common values
        public static bool CheckIsValidExtension(string fileExtension)
        {
            ArrayList arrExtension = new ArrayList();
            arrExtension.Add(".JPG");
            arrExtension.Add(".JPEG");
            arrExtension.Add(".GIF");
            arrExtension.Add(".PNG");
            arrExtension.Add(".BMP");
            arrExtension.Add(".DOC");
            arrExtension.Add(".DOCX");
            arrExtension.Add(".PDF");
            arrExtension.Add(".MP4");
            arrExtension.Add(".AVI");
            arrExtension.Add(".WMV");
            arrExtension.Add(".MKV");
            return arrExtension.Contains(fileExtension.ToUpper());
        }
        public static bool CheckIsValidDocOrImageExtension(string fileExtension)
        {
            ArrayList arrExtension = new ArrayList();
            arrExtension.Add(".JPG");
            arrExtension.Add(".JPEG");
            arrExtension.Add(".GIF");
            arrExtension.Add(".PNG");
            arrExtension.Add(".BMP");
            arrExtension.Add(".DOC");
            arrExtension.Add(".DOCX");
            arrExtension.Add(".PDF");
            return arrExtension.Contains(fileExtension.ToUpper());
        }
        public static bool CheckIsImageExtension(string fileExtension)
        {
            ArrayList arrExtension = new ArrayList();
            arrExtension.Add(".JPG");
            arrExtension.Add(".JPEG");
            arrExtension.Add(".GIF");
            arrExtension.Add(".PNG");
            arrExtension.Add(".BMP");
            return arrExtension.Contains(fileExtension.ToUpper());
        }
        public static bool CheckIsDocumentExtension(string fileExtension)
        {
            ArrayList arrExtension = new ArrayList();
            arrExtension.Add(".DOC");
            arrExtension.Add(".DOCX");
            arrExtension.Add(".PDF");
            return arrExtension.Contains(fileExtension.ToUpper());
        }
        public static bool CheckIsVideoExtension(string fileExtension)
        {
            ArrayList arrExtension = new ArrayList();
            arrExtension.Add(".MP4");
            arrExtension.Add(".AVI");
            arrExtension.Add(".WMV");
            arrExtension.Add(".MKV");
            return arrExtension.Contains(fileExtension.ToUpper());
        }

        public static string GetMonthNameBanglaByShortMonthName(string srtMonthName)
        {
            string monthName = "";
            switch (srtMonthName.ToUpper())
            {
                case "JAN":
                    monthName = "জানুয়ারি";
                    break;
                case "FEB":
                    monthName = "ফেব্রুয়ারী ";
                    break;
                case "MAR":
                    monthName = "মার্চ";
                    break;
                case "APR":
                    monthName = "এপ্রিল";
                    break;
                case "MAY":
                    monthName = "মে";
                    break;
                case "JUN":
                    monthName = "জুন";
                    break;
                case "JUL":
                    monthName = "জুলাই";
                    break;
                case "AUG":
                    monthName = "অগাষ্ট";
                    break;
                case "SEP":
                    monthName = "সেপ্টেম্বর";
                    break;
                case "OCT":
                    monthName = "অক্টোবর";
                    break;
                case "NOV":
                    monthName = "নভেম্বর";
                    break;
                case "DEC":
                    monthName = "ডিসেম্বর";
                    break;
            }
            return monthName;
        }
        public static string GetMonthNameByMonthNumber(short monthNo)
        {
            string monthName = "";
            switch (monthNo)
            {
                case 1:
                    monthName = "January";
                    break;
                case 2:
                    monthName = "February";
                    break;
                case 3:
                    monthName = "March";
                    break;
                case 4:
                    monthName = "April";
                    break;
                case 5:
                    monthName = "May";
                    break;
                case 6:
                    monthName = "June";
                    break;
                case 7:
                    monthName = "July";
                    break;
                case 8:
                    monthName = "August";
                    break;
                case 9:
                    monthName = "September";
                    break;
                case 10:
                    monthName = "October";
                    break;
                case 11:
                    monthName = "November";
                    break;
                case 12:
                    monthName = "December";
                    break;
            }
            return monthName;
        }




        public static string SentenceCase(string input)
        {
            if (input.Length < 1)
                return input;

            string sentence = input.ToLower();
            return sentence[0].ToString().ToUpper() +
               sentence.Substring(1);
        }
        public static short GetLengthFromTwoDate(DateTime fromDate, DateTime toDate)
        {
            short length = 0;
            if (toDate >= fromDate)
            {
                TimeSpan diffDate = toDate.Subtract(fromDate);
                length = Convert.ToInt16(diffDate.Days + 1);
            }
            return length;
        }
        
        public static string ParseText(string inputValue)
        {
            string outputValue = inputValue.Replace("'", "''");
            outputValue = outputValue.Replace("&", "'||'&'||'");
            return outputValue;
        }
        public static DateTime GetFormattedDateMMDDYYYY(string ddMMYYYY)
        {
            string strDateMMDDYY = "";
            if (ddMMYYYY.Contains("-"))
            {
                string[] dateParts = ddMMYYYY.Split('-');
                strDateMMDDYY = dateParts[1] + "-" + dateParts[0] + "-" + dateParts[2] + " " + DateTime.Now.TimeOfDay;
            }
            else
            {
                string[] dateParts = ddMMYYYY.Split('/');
                strDateMMDDYY = dateParts[1] + "/" + dateParts[0] + "/" + dateParts[2] + " " + DateTime.Now.TimeOfDay;
            }
            DateTime MMDDYYYY = Convert.ToDateTime(strDateMMDDYY);
            return MMDDYYYY;
        }
        public static DateTime GetFormattedDateTimeMMDDYYYHM(string ddMMYYYYHM)
        {
            string[] dateTimeParts = ddMMYYYYHM.Split(' ');
            string[] dateParts = dateTimeParts[0].Split('/');
            string strDateMMDDYY = dateParts[1] + "/" + dateParts[0] + "/" + dateParts[2] + " " + dateTimeParts[1] + " " + dateTimeParts[2]; //+ DateTime.Now.TimeOfDay;
                                                                                                                                             //string strDateMMDDYY = dateTimeParts[1] + dateTimeParts[2] + dateTimeParts[0] + dateTimeParts[3] + " " + dateTimeParts[4];// +dateTimeParts[5];
            DateTime MMDDYYYY = Convert.ToDateTime(strDateMMDDYY);
            return MMDDYYYY;
        }
        public static DateTime GetFormattedDateTimeMMDDYYYHMS(string ddMMYYYYHMS)
        {
            string[] dateTimeParts = ddMMYYYYHMS.Split(' ');
            //string[] dateParts = dateTimeParts[0].Split('/');
            //string strDateMMDDYY = dateParts[1] + "/" + dateParts[0] + "/" + dateParts[2] + " " + dateTimeParts[1]; //+ DateTime.Now.TimeOfDay;
            string strDateMMDDYY = dateTimeParts[1] + dateTimeParts[2] + dateTimeParts[0] + dateTimeParts[3] + " " + dateTimeParts[4] + dateTimeParts[5];
            DateTime MMDDYYYY = Convert.ToDateTime(strDateMMDDYY);
            return MMDDYYYY;
        }
        public static DateTime GetStandardFormattedDate(string ddMMyy)
        {
            IFormatProvider culture = new CultureInfo("en-US", true);
            DateTime dt = new DateTime();
            dt = DateTime.ParseExact(ddMMyy, "dd/MM/yy", culture, DateTimeStyles.NoCurrentDateDefault);
            return dt;
        }
        public static DateTime GetDateddMMyyyy(string date){

            // https://www.dotnetperls.com/datetime-parse   
            DateTime dt = DateTime.ParseExact(date.ToString(), "MM/dd/yyyy hh:mm:ss tt", CultureInfo.InvariantCulture);

            string s = dt.ToString("dd/M/yyyy", CultureInfo.InvariantCulture);
            return dt;

        }


        


        

    }
}
