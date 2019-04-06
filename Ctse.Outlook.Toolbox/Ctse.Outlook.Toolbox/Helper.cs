using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Reflection;
//using Microsoft.Office.Interop.Outlook;
using MSO = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Ctse.Outlook.Toolbox.MAPIMessageConverter;

namespace Ctse.Outlook.Toolbox
{
    public class Helper
    {
        //imagemso
        // https://docs.microsoft.com/en-us/openspecs/office_standards/ms-customui/fe2124a1-5aaa-4adf-b285-5d58da9d5e2a
        // https://bert-toolkit.com/imagemso-list.html

        public static void SaveMessageToEml(MSO.MailItem mailItem, bool asTemplate = true)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "*.eml|*.eml";
            saveFileDialog.DefaultExt = "eml";
            if (saveFileDialog.ShowDialog() != DialogResult.OK)
                return;
            string fileName = saveFileDialog.FileName;

            mailItem.Save();
            string entryId = mailItem.EntryID;
            string tempFileName = (asTemplate)?Path.GetTempFileName():fileName;

            Type converter = Type.GetTypeFromCLSID(MAPIMethods.CLSID_IConverterSession);
            object obj = Activator.CreateInstance(converter);
            MAPIMethods.IConverterSession session = (MAPIMethods.IConverterSession)obj;

            if (session != null)
            {
                uint hr = session.SetEncoding(MAPIMethods.ENCODINGTYPE.IET_QP);
                hr = session.SetSaveFormat(MAPIMethods.MIMESAVETYPE.SAVE_RFC1521);
                var stream = new ComMemoryStream();
                hr = session.MAPIToMIMEStm((MAPIMethods.IMessage)mailItem.MAPIOBJECT, stream, MAPIMethods.MAPITOMIMEFLAGS.CCSF_SMTP);
                if (hr != 0)
                    throw new ArgumentException("There are some invalid COM arguments");       

                stream.Position = 0;

                using (FileStream file = new FileStream(tempFileName, FileMode.Create, System.IO.FileAccess.Write))
                {
                    stream.CopyTo(file);
                }

                if (asTemplate)
                {
                    // remove some header-info
                    Encoding encoding = Encoding.Default;
                    string[] skipLines = new string[] { "from:", "to:", "thread-index:", "message-id:", "references:", "in-reply-to:", "date:", "organization:", "x-originating-ip:", "x-unsent:" };
                    var list = File.ReadAllLines(tempFileName, encoding).Where(line =>
                    {
                        return !skipLines.Any(sk => line.ToLower().StartsWith(sk));
                    }).ToList();

                    var idx = list.FindIndex(0, x => x.ToLower().StartsWith("x-"));
                    if (idx >= 0)
                    {
                        list.Insert(idx, "X-Unsent: 1");
                    }
                    else
                    {
                        idx = list.FindIndex(0, m => m.ToLower().StartsWith("mime-version"));
                        if (idx >= 0)
                        {
                            list.Insert(idx, "X-Unsent: 1");
                        }
                    }
                    File.WriteAllLines(fileName, list, encoding);
                    try
                    {
                        File.Delete(tempFileName);
                    }
                    catch
                    {
                    }
                }
            }

            ComSafeHelper.Release(mailItem);

        }
    }
}
