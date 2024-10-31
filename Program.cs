using System;
using System.Windows.Forms;
using MicroLabPC;

namespace WindowsFormsApp1
{
    internal static class Program
    { 
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try
            {
                int init = MLInstrumentInvoke.Init();

                if (init == 0)
                {
                    Console.WriteLine("Initialization successful: " + init);
                }
                else
                {
                    Console.WriteLine("Initialization failed with error code: " + init);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception during initialization: " + ex.Message);
            }
            Application.Run();
 
        }
    }
}
