using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

using GlobalStructures;

namespace WinUI3_SendMail
{
    internal class MAPI
    {
        [DllImport("MAPI32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern HRESULT MAPIInitialize(IntPtr lpMapiInit);

        internal delegate MAPI_ERROR MAPISendMailDelegate(IntPtr lhSession, UIntPtr ulUIParam, IntPtr lpMessage, int flFlags, uint ulReserved);

        //[DllImport("hhmapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        //public static extern MAPI_ERROR MAPISendMail(IntPtr lhSession, UIntPtr ulUIParam, IntPtr lpMessage, int flFlags, uint ulReserved);


        public static uint PROP_TAG(uint ulPropType, uint ulPropID) { return ((((ulPropID)) << 16) | ((ulPropType))); }

        public const int MV_FLAG       =  0x1000         ;  /* Multi-value flag */

        public const uint PT_UNSPECIFIED = ((uint)0);  /* (Reserved for interface use) type doesn't matter to caller */
        public const uint PT_NULL = ((uint)1);  /* NULL property value */
        public const uint PT_I2 = ((uint)2);  /* Signed 16-bit value */
        public const uint PT_LONG = ((uint)3);  /* Signed 32-bit value */
        public const uint PT_R4 = ((uint)4);  /* 4-byte floating point */
        public const uint PT_DOUBLE = ((uint)5);  /* Floating point double */
        public const uint PT_CURRENCY = ((uint)6);  /* Signed 64-bit int (decimal w/    4 digits right of decimal pt) */
        public const uint PT_APPTIME = ((uint)7);  /* Application time */
        public const uint PT_ERROR = ((uint)10);  /* 32-bit error value */
        public const uint PT_BOOLEAN = ((uint)11);  /* 16-bit boolean (non-zero true) */
        public const uint PT_OBJECT = ((uint)13);  /* Embedded object in a property */
        public const uint PT_I8 = ((uint)20);  /* 8-byte signed integer */
        public const uint PT_STRING8 = ((uint)30);  /* Null terminated 8-bit character string */
        public const uint PT_UNICODE = ((uint)31);  /* Null terminated Unicode string */
        public const uint PT_SYSTIME = ((uint)64);  /* FILETIME 64-bit int w/ number of 100ns periods since Jan 1,1601 */
        public const uint PT_CLSID = ((uint)72);  /* OLE GUID */
        public const uint PT_BINARY = ((uint)258);  /* Uninterpreted (counted byte array) */
        /* Changes are likely to these numbers, and to their structures. */

        /* Alternate property type names for ease of use */
        public const uint PT_SHORT = PT_I2;
        public const uint PT_I4 = PT_LONG;
        public const uint PT_FLOAT = PT_R4;
        public const uint PT_R8 = PT_DOUBLE;
        public const uint PT_LONGLONG = PT_I8;

        public const uint PT_TSTRING = PT_UNICODE;
        public const uint PT_MV_TSTRING = (MV_FLAG | PT_UNICODE);

        public static uint PR_DISPLAY_NAME = PROP_TAG(PT_TSTRING, 0x3001);

        public const int MAPI_UNREAD = 0x00000001;
        public const int MAPI_RECEIPT_REQUESTED = 0x00000002;
        public const int MAPI_SENT = 0x00000004;

        /* MAPILogon() flags.       */
        public const int MAPI_LOGON_UI = 0x00000001;  /* Display logon UI             */
        public const int MAPI_PASSWORD_UI = 0x00020000; /* prompt for password only     */
        public const int MAPI_NEW_SESSION = 0x00000002;  /* Don't use shared session     */
        public const int MAPI_FORCE_DOWNLOAD = 0x00001000;  /* Get new mail before return   */
        public const int MAPI_EXTENDED = 0x00000020;  /* Extended MAPI Logon          */

        /* MAPISendMail() flags.    */
        /* also defined in property.h */
        public const int MAPI_DIALOG = 0x00000008;  /* Display a send note UI       */
        public const int MAPI_USE_DEFAULT = 0x00000040; /* Use default profile in logon */

        /* MAPISendMailW() flags.    */
        public const int MAPI_DIALOG_MODELESS = 0x00000004 | MAPI_DIALOG;  /* Display a modeless window    */
        public const int MAPI_FORCE_UNICODE = 0x00040000;  /* Don't down-convert to ANSI if provider does not support Unicode */

        /* MAPIFindNext() flags.    */
        public const int MAPI_UNREAD_ONLY = 0x00000020;  /* Only unread messages         */
        public const int MAPI_GUARANTEE_FIFO = 0x00000100;  /* use date order               */
        public const int MAPI_LONG_MSGID = 0x00004000;  /* allow 512 char returned ID	*/

        /* MAPIReadMail() flags.    */
        public const int MAPI_PEEK = 0x00000080;  /* Do not mark as read.         */
        public const int MAPI_SUPPRESS_ATTACH = 0x00000800;  /* header + body, no files      */
        public const int MAPI_ENVELOPE_ONLY = 0x00000040;  /* Only header information      */
        public const int MAPI_BODY_AS_FILE = 0x00000200;

        /* MAPISaveMail() flags.    */
        /* #define MAPI_LOGON_UI        0x00000001     Display logon UI             */
        /* #define MAPI_NEW_SESSION     0x00000002     Don't use shared session     */
        /* #define MAPI_LONG_MSGID		0x00004000	/* allow 512 char returned ID	*/

        /* MAPIAddress() flags.     */
        /* #define MAPI_LOGON_UI        0x00000001     Display logon UI             */
        /* #define MAPI_NEW_SESSION     0x00000002     Don't use shared session     */

        /* MAPIDetails() flags.     */
        /* #define MAPI_LOGON_UI        0x00000001     Display logon UI             */
        /* #define MAPI_NEW_SESSION     0x00000002     Don't use shared session     */
        public const int MAPI_AB_NOMODIFY = 0x00000400;  /* Don't allow mods of AB entries */

        public enum MAPI_ERROR : uint
        {
            SUCCESS_SUCCESS = 0,
            MAPI_USER_ABORT = 1,
            MAPI_E_USER_ABORT = MAPI_USER_ABORT,
            MAPI_E_FAILURE = 2,
            MAPI_E_LOGON_FAILURE = 3,
            MAPI_E_LOGIN_FAILURE = MAPI_E_LOGON_FAILURE,
            MAPI_E_DISK_FULL = 4,
            MAPI_E_INSUFFICIENT_MEMORY = 5,
            MAPI_E_ACCESS_DENIED = 6,
            MAPI_E_TOO_MANY_SESSIONS = 8,
            MAPI_E_TOO_MANY_FILES = 9,
            MAPI_E_TOO_MANY_RECIPIENTS = 10,
            MAPI_E_ATTACHMENT_NOT_FOUND = 11,
            MAPI_E_ATTACHMENT_OPEN_FAILURE = 12,
            MAPI_E_ATTACHMENT_WRITE_FAILURE = 13,
            MAPI_E_UNKNOWN_RECIPIENT = 14,
            MAPI_E_BAD_RECIPTYPE = 15,
            MAPI_E_NO_MESSAGES = 16,
            MAPI_E_INVALID_MESSAGE = 17,
            MAPI_E_TEXT_TOO_LARGE = 18,
            MAPI_E_INVALID_SESSION = 19,
            MAPI_E_TYPE_NOT_SUPPORTED = 20,
            MAPI_E_AMBIGUOUS_RECIPIENT = 21,
            MAPI_E_AMBIG_RECIP = MAPI_E_AMBIGUOUS_RECIPIENT,
            MAPI_E_MESSAGE_IN_USE = 22,
            MAPI_E_NETWORK_FAILURE = 23,
            MAPI_E_INVALID_EDITFIELDS = 24,
            MAPI_E_INVALID_RECIPS = 25,
            MAPI_E_NOT_SUPPORTED = 26,
            MAPI_E_UNICODE_NOT_SUPPORTED = 27,
            MAPI_E_ATTACHMENT_TOO_LARGE = 28
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        public struct MapiMessage
        {
            public uint ulReserved;             /* Reserved for future use (M.B. 0)       */
            public string lpszSubject;            /* Message Subject                        */
            public string lpszNoteText;           /* Message Text                           */
            public string lpszMessageType;        /* Message Class                          */
            public string lpszDateReceived;       /* in YYYY/MM/DD HH:MM format             */
            public string lpszConversationID;     /* conversation thread ID                 */
            public int flFlags;                /* unread,return receipt                  */
            public IntPtr lpOriginator; /* Originator descriptor                  */
            public uint nRecipCount;            /* Number of recipients                   */
            public IntPtr lpRecips;     /* Recipient descriptors                  */
            public uint nFileCount;             /* # of file attachments                  */
            public IntPtr lpFiles;       /* Attachment descriptors                 */
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct MapiMessageW
        {
            public uint ulReserved;
            public string lpszSubject;
            public string lpszNoteText;
            public string lpszMessageType;
            public string lpszDateReceived;
            public string lpszConversationID;
            public int flFlags;
            //public lpMapiRecipDescW lpOriginator;
            public IntPtr lpOriginator;
            public uint nRecipCount;
            //public lpMapiRecipDescW lpRecips;
            public IntPtr lpRecips;
            public uint nFileCount;
            //public lpMapiFileDescW lpFiles;
            public IntPtr lpFiles;
        };

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct MapiRecipDescW
        {
            public uint ulReserved;
            public uint ulRecipClass;
            public string lpszName;
            public string lpszAddress;
            public uint ulEIDSize;
            public IntPtr lpEntryID;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        public struct MapiRecipDesc
        {
            public uint ulReserved;
            public uint ulRecipClass;
            public string lpszName;
            public string lpszAddress;
            public uint ulEIDSize;
            public IntPtr lpEntryID;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        public struct MapiFileDesc
        {
            public uint ulReserved;            /* Reserved for future use (must be 0)     */
            public uint flFlags;               /* Flags                                   */
            public uint nPosition;             /* character in text to be replaced by attachment */
            public string lpszPathName;          /* Full path name of attachment file       */
            public string lpszFileName;          /* Original file name (optional)           */
            public IntPtr lpFileType;           /* Attachment file type (can be lpMapiFileTagExt) */

            public MapiFileDesc(uint ulReserved, uint flFlags, uint nPosition, string lpszPathName, string lpszFileName, IntPtr lpFileType)
            {
                this.ulReserved = ulReserved;
                this.flFlags = flFlags;
                this.nPosition = nPosition;
                this.lpszPathName = lpszPathName;
                this.lpszFileName = lpszFileName;
                this.lpFileType = lpFileType;
            }
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct MapiFileDescW
        {
            public uint ulReserved;
            public uint flFlags;
            public uint nPosition;
            public string lpszPathName;
            public string lpszFileName;
            public IntPtr lpFileType;

            public MapiFileDescW(uint ulReserved, uint flFlags, uint nPosition, string lpszPathName, string lpszFileName, IntPtr lpFileType)
            {
                this.ulReserved = ulReserved;
                this.flFlags = flFlags;
                this.nPosition = nPosition;
                this.lpszPathName = lpszPathName;
                this.lpszFileName = lpszFileName;
                this.lpFileType = lpFileType;
            }
        }

        public const int MAPI_ORIG = 0;          /* Recipient is message originator          */
        public const int MAPI_TO = 1;          /* Recipient is a primary recipient         */
        public const int MAPI_CC = 2;         /* Recipient is a copy recipient            */
        public const int MAPI_BCC = 3;          /* Recipient is blind copy recipient        */

       
    }
}
