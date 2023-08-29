using AventStack.ExtentReports.Model;
using ExcelMonitor.Properties;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using Excel = Microsoft.Office.Interop.Excel;
using Rectangle = System.Drawing.Rectangle;

namespace ExcelMonitor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            Text = Settings.Default.zalozka;
            GetImage();
            System.Timers.Timer aTimer = new System.Timers.Timer();
            aTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
            aTimer.Interval = Convert.ToInt32(Settings.Default.frekvencia);
            aTimer.Enabled = true;
        }

        void GetImage()
        {
            if (pictureBox1.Image != null)
            {
                pictureBox1.Image.Dispose();
                pictureBox1.Image = null;
            }
                Exception threadEx = null;
                Thread staThread = new Thread(
                    delegate ()
                    {
                        try
                        {
                            var a = new Application();
                            Workbook w = a.Workbooks.Open(@Settings.Default.adresa.ToString(), false);
                            Worksheet ws = w.Sheets[Settings.Default.zalozka];
                            Range r = ws.Range[Settings.Default.bunky];
                            try
                            {
                            Clipboard.Clear();
                            r.CopyPicture(XlPictureAppearance.xlScreen, XlCopyPictureFormat.xlBitmap); 
                            }
                            catch (Exception)
                            {
                                Marshal.ReleaseComObject(r);
                                Marshal.ReleaseComObject(ws);
                                //w.Save();
                                w.Close(false);
                                Marshal.ReleaseComObject(w);
                                a.Quit();
                                Marshal.ReleaseComObject(a);

                                GC.Collect();
                                GC.WaitForPendingFinalizers();
                                return; 
                            }

                            Bitmap image = new Bitmap(Clipboard.GetImage());
                            pictureBox1.Image = image;

                            Marshal.ReleaseComObject(r);
                            Marshal.ReleaseComObject(ws);
                            //w.Save();
                            w.Close(false);
                            Marshal.ReleaseComObject(w);
                            a.Quit();
                            Marshal.ReleaseComObject(a);

                            GC.Collect();
                            GC.WaitForPendingFinalizers();

                        }
                        catch (Exception ex)
                        {
                             threadEx = ex;
                        }
                    });
                staThread.SetApartmentState(ApartmentState.STA);
                staThread.Start();
                staThread.Join();
        }

        void OnTimedEvent(object source, ElapsedEventArgs e)
        {
            GetImage();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            timer1.Stop();
        }

    }
}
