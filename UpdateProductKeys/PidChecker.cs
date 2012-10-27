using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace UpdateProductKeys
{
    class PidChecker
    {
        List<string> pidsChecked = new List<string>();
        List<string> pidsBad = new List<string>();
        bool pracDB;

        [DllImport("pidgenx.dll", EntryPoint = "PidGenX", CharSet = CharSet.Auto)]
        static extern int PidGenX(string ProductKey, string PkeyPath, string MSPID, int UnknownUsage, IntPtr ProductID, IntPtr DigitalProductID, IntPtr DigitalProductID4);

        internal PidChecker(string production)
        {
            if (production != "-p")
            {
                Console.WriteLine("Using Practice Database");
                pracDB = true;
            }
            else
                Console.WriteLine("Using Production Database");
        }

        internal string CheckProductKey(string productKey)
        {
            //set a conditional compilation symbol to avoid delays of pid checking while in debug just put pidNo in conditional compilation symbols in properties
#if pidNo
          return "Valid" 
#endif
            if (pracDB == true)
            {// do not do pid chekcing on practice db
                return "Valid";
            }
            if (pidsBad.Contains(productKey))
                return "Invalid Product Key - see message for this key previously";
            if (pidsChecked.Contains(productKey))
                return "Valid";
            string result = "";
            int RetID;
            byte[] gpid = new byte[0x32];
            byte[] opid = new byte[0xA4];
            byte[] npid = new byte[0x04F8];

            IntPtr PID = Marshal.AllocHGlobal(0x32);
            IntPtr DPID = Marshal.AllocHGlobal(0xA4);
            IntPtr DPID4 = Marshal.AllocHGlobal(0x04F8);

            string PKeyPath = Environment.SystemDirectory + @"\spp\tokens\pkeyconfig\pkeyconfig.xrm-ms";
            string MSPID = "00000";

            gpid[0] = 0x32;
            opid[0] = 0xA4;
            npid[0] = 0xF8;
            npid[1] = 0x04;

            Marshal.Copy(gpid, 0, PID, 0x32);
            Marshal.Copy(opid, 0, DPID, 0xA4);
            Marshal.Copy(npid, 0, DPID4, 0x04F8);

            RetID = PidGenX(productKey, PKeyPath, MSPID, 0, PID, DPID, DPID4);

            if (RetID == 0)
            {
                Marshal.Copy(PID, gpid, 0, gpid.Length);
                Marshal.Copy(DPID4, npid, 0, npid.Length);
                string edi = GetString(npid, 0x0118);
                string lit = GetString(npid, 0x03F8);
                Console.WriteLine("edi = " + edi);
                Console.WriteLine("lit = " + lit);
                if (!(edi == "Professional" || edi == ""))
                {
                    result = "This is not a Windows 7 Professional Product Key";
                    pidsBad.Add(productKey);
                }
                else if (lit == "OEM:SLP" || lit == "OEM:NONSLP")
                {
                    result = "This Key type is not allowed. The OEM:COA key off the machine label must be used";
                    pidsBad.Add(productKey);
                }
                else
                {
                    result = "Valid";
                    pidsChecked.Add(productKey);
                }
            }
            else if (RetID == -2147024809)
            {
                result = "PIDChecker - Invalid Arguments";
                pidsBad.Add(productKey);
            }
            else if (RetID == -1979645695)
            {
                result = "Not a Windows 7 Product Key";
                pidsBad.Add(productKey);
            }
            else if (RetID == -2147024894)
            {
                result = "PidChecker - pkeyconfig.xrm.ms file is not found";
                pidsBad.Add(productKey);
            }
            else
            {
                result = "PidChecker - Invalid input!!!";
                pidsBad.Add(productKey);
            }
            Marshal.FreeHGlobal(PID);
            Marshal.FreeHGlobal(DPID);
            Marshal.FreeHGlobal(DPID4);
            //FreeLibrary(dllHandle);
            return result;
        }
        string GetString(byte[] bytes, int index)
        {
            int n = index;
            while (!(bytes[n] == 0 && bytes[n + 1] == 0)) n++;
            return Encoding.ASCII.GetString(bytes, index, n - index).Replace("\0", "");
        }
    }
}