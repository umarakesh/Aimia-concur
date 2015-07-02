using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace BOS_PO_FROM_CSV_ConCur
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            pofromcsv_cls oHelloWorld = new pofromcsv_cls();
            Global.globaltime1 = DateTime.Now;
            Global.globaltime = DateTime.Now.ToString("yyMMddHHmmss");
            System.Windows.Forms.Application.Run();
        }
    }
}
