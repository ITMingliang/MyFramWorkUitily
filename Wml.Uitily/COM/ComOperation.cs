using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Wml.Uitily.COM
{
    class ComOperation
    {

        #region 结构体和dllimport
        //public string PortNum;   //1,2,3,4 
        //public int BaudRate;   //1200,2400,4800,9600 
        //public byte ByteSize;   //8   bits 
        //public byte Parity;   //   0-4=no,odd,even,mark,space   
        //public byte StopBits;   //   0,1,2   =   1,   1.5,   2   
        //public int ReadTimeout;   //10 

        //设置端口
        public string PortNum = "COM1";   //端口 1,2,3,4  

        //端口设置(默认值，需根据端口情况重新赋值)
        public int BaudRate = 2400;    //1200,2400,4800,9600 每秒位数
        public byte ByteSize = 8;      //8   bits  数据位
        public byte Parity = 0;        //   0-4=no,odd,even,mark,space   奇偶效验
        public byte StopBits = 0;      //   0,1,2   =   1,   1.5,   2   停止位
        public int ReadTimeout = 10;   //10 

        //comm   port   win32   file   handle 
        private int hComm = -1;

        public bool Opened = false;

        //win32   api   constants 
        private const uint GENERIC_READ = 0x80000000;
        private const uint GENERIC_WRITE = 0x40000000;
        private const int OPEN_EXISTING = 3;
        private const int INVALID_HANDLE_VALUE = -1;

        [StructLayout(LayoutKind.Sequential)]
        private struct DCB
        {
            //taken   from   c   struct   in   platform   sdk   
            public int DCBlength;                       //   sizeof(DCB)   
            public int BaudRate;                         //   current   baud   rate   
            public int fBinary;                     //   binary   mode,   no   EOF   check   
            public int fParity;                     //   enable   parity   checking   
            public int fOutxCtsFlow;             //   CTS   output   flow   control   
            public int fOutxDsrFlow;             //   DSR   output   flow   control   
            public int fDtrControl;               //   DTR   flow   control   type   
            public int fDsrSensitivity;       //   DSR   sensitivity   
            public int fTXContinueOnXoff;   //   XOFF   continues   Tx   
            public int fOutX;                     //   XON/XOFF   out   flow   control   
            public int fInX;                       //   XON/XOFF   in   flow   control   
            public int fErrorChar;           //   enable   error   replacement   
            public int fNull;                     //   enable   null   stripping   
            public int fRtsControl;           //   RTS   flow   control   
            public int fAbortOnError;       //   abort   on   error   
            public int fDummy2;                 //   reserved   
            public ushort wReserved;                     //   not   currently   used   
            public ushort XonLim;                           //   transmit   XON   threshold   
            public ushort XoffLim;                         //   transmit   XOFF   threshold   
            public byte ByteSize;                       //   number   of   bits/byte,   4-8   
            public byte Parity;                           //   0-4=no,odd,even,mark,space   
            public byte StopBits;                       //   0,1,2   =   1,   1.5,   2   
            public char XonChar;                         //   Tx   and   Rx   XON   character   
            public char XoffChar;                       //   Tx   and   Rx   XOFF   character   
            public char ErrorChar;                     //   error   replacement   character   
            public char EofChar;                         //   end   of   input   character   
            public char EvtChar;                         //   received   event   character   
            public ushort wReserved1;                   //   reserved;   do   not   use   
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct COMMTIMEOUTS
        {
            public int ReadIntervalTimeout;
            public int ReadTotalTimeoutMultiplier;
            public int ReadTotalTimeoutConstant;
            public int WriteTotalTimeoutMultiplier;
            public int WriteTotalTimeoutConstant;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct OVERLAPPED
        {
            public int Internal;
            public int InternalHigh;
            public int Offset;
            public int OffsetHigh;
            public int hEvent;
        }

        [DllImport("kernel32.dll ")]
        private static extern int CreateFile(
        string lpFileName,                                                   //   file   name 
        uint dwDesiredAccess,                                             //   access   mode 
        int dwShareMode,                                                     //   share   mode 
        int lpSecurityAttributes,   //   SD 
        int dwCreationDisposition,                                 //   how   to   create 
        int dwFlagsAndAttributes,                                   //   file   attributes 
        int hTemplateFile                                                 //   handle   to   template   file 
        );
        [DllImport("kernel32.dll ")]
        private static extern bool GetCommState(
        int hFile,     //   handle   to   communications   device 
        ref DCB lpDCB         //   device-control   block 
        );
        [DllImport("kernel32.dll ")]
        private static extern bool BuildCommDCB(
        string lpDef,     //   device-control   string 
        ref DCB lpDCB           //   device-control   block 
        );
        [DllImport("kernel32.dll ")]
        private static extern bool SetCommState(
        int hFile,     //   handle   to   communications   device 
        ref DCB lpDCB         //   device-control   block 
        );
        [DllImport("kernel32.dll ")]
        private static extern bool GetCommTimeouts(
        int hFile,                                     //   handle   to   comm   device 
        ref COMMTIMEOUTS lpCommTimeouts     //   time-out   values 
        );
        [DllImport("kernel32.dll ")]
        private static extern bool SetCommTimeouts(
        int hFile,                                     //   handle   to   comm   device 
        ref COMMTIMEOUTS lpCommTimeouts     //   time-out   values 
        );
        [DllImport("kernel32.dll ")]
        private static extern bool ReadFile(
        int hFile,                                 //   handle   to   file 
        byte[] lpBuffer,                           //   data   buffer 
        int nNumberOfBytesToRead,     //   number   of   bytes   to   read 
        ref int lpNumberOfBytesRead,   //   number   of   bytes   read 
        ref OVERLAPPED lpOverlapped         //   overlapped   buffer 
        );
        [DllImport("kernel32.dll ")]
        private static extern bool WriteFile(
        int hFile,                                         //   handle   to   file 
        byte[] lpBuffer,                                 //   data   buffer 
        int nNumberOfBytesToWrite,           //   number   of   bytes   to   write 
        ref int lpNumberOfBytesWritten,     //   number   of   bytes   written 
        ref OVERLAPPED lpOverlapped                 //   overlapped   buffer 
        );
        [DllImport("kernel32.dll ")]
        private static extern bool CloseHandle(
        int hObject       //   handle   to   object 
        );
        #endregion


        #region 设置端口参数
        /// <summary>
        /// 设置端口参数
        /// </summary>
        /// <param name="parPortNum"></param>
        /// <param name="parBaudRate"></param>
        /// <param name="parByteSize"></param>
        /// <param name="parParity"></param>
        /// <param name="parStopBits"></param>
        public void SetComPara(string parPortNum, int parBaudRate, byte parByteSize, byte parParity, byte parStopBits)
        {
            //设置端口
            this.PortNum = parPortNum;    //端口 1,2,3,4  

            this.BaudRate = parBaudRate;    //1200,2400,4800,9600 每秒位数
            this.ByteSize = parByteSize;      //8   bits  数据位
            this.Parity = parParity;        //   0-4=no,odd,even,mark,space   奇偶效验
            this.StopBits = parStopBits;      //   0,1,2   =   1,   1.5,   2   停止位      
        }
        #endregion

        #region 打开Com端口
        /// <summary>
        /// 打开Com端口
        /// </summary>
        public void OpenCommPort()
        {

            try
            {
                DCB dcbCommPort = new DCB();
                COMMTIMEOUTS ctoCommPort = new COMMTIMEOUTS();


                //   OPEN   THE   COMM   PORT. 
                hComm = CreateFile(PortNum, GENERIC_READ | GENERIC_WRITE, 0, 0, OPEN_EXISTING, 0, 0);

                //   IF   THE   PORT   CANNOT   BE   OPENED,   BAIL   OUT. 
                if (hComm == INVALID_HANDLE_VALUE)
                {
                    throw (new ApplicationException("Comm   Port   Can   Not   Be   Opened "));
                }

                //   SET   THE   COMM   TIMEOUTS. 
                GetCommTimeouts(hComm, ref ctoCommPort);
                ctoCommPort.ReadTotalTimeoutConstant = ReadTimeout;
                ctoCommPort.ReadTotalTimeoutMultiplier = 0;
                ctoCommPort.WriteTotalTimeoutMultiplier = 0;
                ctoCommPort.WriteTotalTimeoutConstant = 0;
                SetCommTimeouts(hComm, ref ctoCommPort);

                //   SET   BAUD   RATE,   PARITY,   WORD   SIZE,   AND   STOP   BITS. 
                //   THERE   ARE   OTHER   WAYS   OF   DOING   SETTING   THESE   BUT   THIS   IS   THE   EASIEST. 
                //   IF   YOU   WANT   TO   LATER   ADD   CODE   FOR   OTHER   BAUD   RATES,   REMEMBER 
                //   THAT   THE   ARGUMENT   FOR   BuildCommDCB   MUST   BE   A   POINTER   TO   A   STRING. 
                //   ALSO   NOTE   THAT   BuildCommDCB()   DEFAULTS   TO   NO   HANDSHAKING. 

                dcbCommPort.DCBlength = Marshal.SizeOf(dcbCommPort);
                GetCommState(hComm, ref dcbCommPort);

                //将设置的端口参数赋给DCB结构体
                dcbCommPort.BaudRate = BaudRate;
                dcbCommPort.Parity = Parity;
                dcbCommPort.ByteSize = ByteSize;
                dcbCommPort.StopBits = StopBits;
                SetCommState(hComm, ref dcbCommPort);

                Opened = true;


            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion


        #region 关闭Com端口
        /// <summary>
        /// 关闭Com端口 
        /// </summary>
        public void CloseComPort()
        {
            try
            {
                if (hComm != INVALID_HANDLE_VALUE)
                {
                    CloseHandle(hComm);
                    Opened = false;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion
    }
}
