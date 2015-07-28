using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace Video_exporter
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
           // Application.ActivePresentation.CreateVideo();
            export();
            
        }

        public void export()
        {
            Globals.Ribbons.GetRibbon<Video_export>().button1.Click+=button1_Click;
            
        }
        Video_export v;
        void button1_Click(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
        {
            v = Globals.Ribbons.GetRibbon<Video_export>();
            Application.ActivePresentation.CreateVideo(Application.ActivePresentation.FullName + ".mp4", false, Convert.ToInt32(v.editBox1.Text), Convert.ToInt32(v.editBox2.Text), Convert.ToInt32(v.editBox3.Text), 85);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }



        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
