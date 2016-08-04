using Extensibility;
using Microsoft.Office.Core;
using System;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows;
//using Application = Microsoft.Office.Interop.OneNote.Application;

namespace RichPaste
{
    [Guid("C06F8EC7-88DD-4D36-A24D-0EE91CB22B4E")]
    [ProgId("RichPaste.Paste")]
    [Serializable()]
    public class Paste : IDTExtensibility2, IRibbonExtensibility
    {
        //Application onApp = new Application();
        

        public IStream GetImage(string RTF)
        {
            MemoryStream mem = new MemoryStream();
            Properties.Resources.RTF.Save(mem, ImageFormat.Png);
            return new CCOMStreamWrapper(mem);

        }
       
        public string GetCustomUI(string RibbonID)
        {
            return Properties.Resources.ribbon;
        }

        
        public void OnAddInsUpdate(ref Array custom)
        {
        }

        
        public void OnBeginShutdown(ref Array custom)
        {
            //if (onApp != null)
            //onApp = null;

        }

        
        public void OnConnection(object application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
           // onApp = (Application)application;
        }

        
        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            //onApp = null;
            //GC.Collect();
            //GC.WaitForPendingFinalizers();

        }

       
        public void OnStartupComplete(ref Array custom)
        {
        }

        [STAThread]
        public void PasteAction(IRibbonControl control)
        {
            MessageBox.Show("Button Pushed");
            RichConsole.Program.Main();
        }

    }
}
