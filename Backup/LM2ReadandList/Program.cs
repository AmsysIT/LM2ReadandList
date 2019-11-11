using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace LM2ReadandList
{
    static class Program
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Main());
            //Application.Run(new CylinderNoWeight());
        }
    }
}