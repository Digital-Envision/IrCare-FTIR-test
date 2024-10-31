using System;
using System.Runtime.InteropServices;

namespace WindowsFormsApp1.FTIRInterop
{
    public class FTIRInteropMain
    {
        [DllImport("FTIRInst.dll", CallingConvention = CallingConvention.Cdecl)]
        public static extern int Init();

        [DllImport("FTIRInst.dll", EntryPoint = "FTIRInst_SetTargetDeviceUsb", CharSet = CharSet.Unicode)]
        public static extern int FTIRInst_SetTargetDeviceUsb(string pDeviceName);

        [DllImport("ClassLibrary1.dll", EntryPoint = "Add", CharSet = CharSet.Unicode)]
        public static extern int Add(int a, int b);

        public int InitInstrument()
        {
            return Init();
        }

        public int SetTargetDeviceUsb(string pDeviceName)
        {
            Console.WriteLine("Setting target devices...");
            return FTIRInst_SetTargetDeviceUsb(pDeviceName);
        }

        public int AddNumber(int a, int b)
        {
            Console.WriteLine("Adding...");
            return Add(a, b);
        }
    }
}
