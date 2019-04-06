using Microsoft.Office.Interop.Outlook;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace Ctse.Outlook.Toolbox.MAPIMessageConverter
{
    internal class MAPIMethods
    {
        [Flags]
        public enum MAPITOMIMEFLAGS
        {
            CCSF_SMTP = 0x0002,
            CCSF_NOHEADERS = 0x0004,
            CCSF_USE_TNEF = 0x0010,
            CCSF_INCLUDE_BCC = 0x0020,
            CCSF_8BITHEADERS = 0x0040,
            CCSF_USE_RTF = 0x0080,
            CCSF_PLAIN_TEXT_ONLY = 0x1000,
            CCSF_NO_MSGID = 0x4000,
        }

        [Flags]
        public enum CLSCTX
        {
            CLSCTX_INPROC_SERVER = 0x1,
            CLSCTX_INPROC_HANDLER = 0x2,
            CLSCTX_LOCAL_SERVER = 0x4,
            CLSCTX_REMOTE_SERVER = 0x10,
            CLSCTX_INPROC = CLSCTX_INPROC_SERVER | CLSCTX_INPROC_HANDLER,
            CLSCTX_SERVER = CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER | CLSCTX_REMOTE_SERVER,
            CLSCTX_ALL = CLSCTX_SERVER | CLSCTX_INPROC_HANDLER
        }


        public static Guid CLSID_IConverterSession = new Guid("{4e3a7680-b77a-11d0-9da5-00c04fd65685}");

        public static Guid IID_IConverterSession = new Guid("{4b401570-b77b-11d0-9da5-00c04fd65685}");

        public enum ENCODINGTYPE
        {
            IET_BINARY = 0,
            IET_BASE64 = 1,
            IET_UUENCODE = 2,
            IET_QP = 3,
            IET_7BIT = 4,
            IET_8BIT = 5,
            IET_INETCSET = 6,
            IET_UNICODE = 7,
            IET_RFC1522 = 8,
            IET_ENCODED = 9,
            IET_CURRENT = 10,
            IET_UNKNOWN = 11,
            IET_BINHEX40 = 12,
            IET_LAST = 13
        }

        public enum MIMESAVETYPE
        {
            SAVE_RFC822 = 0,
            SAVE_RFC1521 = 1
        }

        [ComVisible(false)]
        [ComImport()]
        [Guid("00020307-0000-0000-C000-000000000046")]
        [InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IMessage
        {
        }

        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("4b401570-b77b-11d0-9da5-00c04fd65685")]
        public interface IConverterSession
        {
            [PreserveSig]
            int Placeholder0();

            [PreserveSig]
            uint SetEncoding(
            [In, MarshalAs(UnmanagedType.I4)] ENCODINGTYPE DispId
            );

            [PreserveSig]
            int Placeholder1();

            [PreserveSig]
            uint MIMEToMAPI(
                [In, MarshalAs(UnmanagedType.Interface)]
                Stream pstm,
                [Out, MarshalAs(UnmanagedType.Interface)]
                MailItem pmsg,
                object pszSrcSrv,
                uint ulFlags
            );

            [PreserveSig]
            uint MAPIToMIMEStm(
                [In, MarshalAs(UnmanagedType.Interface)]
                IMessage pmsg,
                [Out, MarshalAs(UnmanagedType.Interface)]
                IStream pstm,
                MAPITOMIMEFLAGS ulFlags
            );

            [PreserveSig]
            int Placeholder2();

            [PreserveSig]
            int Placeholder3();

            [PreserveSig]
            int Placeholder4();

            [PreserveSig]
            int SetTextWrapping(
                bool fWrapText,
                uint ulWrapWidth
            );

            [PreserveSig]
            uint SetSaveFormat(
                [In, MarshalAs(UnmanagedType.I4)]
                MIMESAVETYPE mstSaveFormat
            );

            [PreserveSig]
            int Placeholder5();

            [PreserveSig]
            int Placeholder6();
        }
    }

    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class ComMemoryStream : MemoryStream, IStream
    {
        public ComMemoryStream() : base() { }

        public ComMemoryStream(byte[] data) : base(data) { }

        #region IStream Members

        public void Clone(out IStream ppstm) { ppstm = null; }

        public void Commit(int grfCommitFlags) { }

        public void CopyTo(
           IStream pstm, long cb, IntPtr pcbRead, IntPtr pcbWritten)
        { }

        public void LockRegion(long libOffset, long cb, int dwLockType) { }

        public void Read(byte[] pv, int cb, IntPtr pcbRead)
        {
            long bytes_read = base.Read(pv, 0, cb);
            if (pcbRead != IntPtr.Zero)
                Marshal.WriteInt64(pcbRead, bytes_read);
        }

        public void Revert() { }

        public void Seek(long dlibMove, int dwOrigin, IntPtr plibNewPosition)
        {
            long pos = base.Seek(dlibMove, (SeekOrigin)dwOrigin);
            if (plibNewPosition != IntPtr.Zero)
                Marshal.WriteInt64(plibNewPosition, pos);
        }

        public void SetSize(long libNewSize) { }

        public void Stat(
           out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg,
           int grfStatFlag)
        {
            pstatstg = new System.Runtime.InteropServices.ComTypes.STATSTG();
            pstatstg.cbSize = base.Length;
        }

        public void UnlockRegion(long libOffset, long cb, int dwLockType) { }

        public void Write(byte[] pv, int cb, IntPtr pcbWritten)
        {
            base.Write(pv, 0, cb);
            if (pcbWritten != IntPtr.Zero)
                Marshal.WriteInt64(pcbWritten, (long)cb);
        }

        #endregion
    }


}