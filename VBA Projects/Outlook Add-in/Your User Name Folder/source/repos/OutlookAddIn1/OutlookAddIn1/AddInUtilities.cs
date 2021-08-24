using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn1
{
    [ComVisible(true)]
    public interface IAddInUtilities
    {
        void ImportData(String xml);
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class AddInUtilities : IAddInUtilities
    {
        // This method tries to writes out
        
        public void ImportData(String ixml)
        {
            // get destination

            // if destination exists write data
            Xml.body = ixml;
        }
    }
    public class Xml
    {
        public static string body = "";

        public static string Body
        {
            get { return body; }
            set { body = value; }
        }
    }
}
