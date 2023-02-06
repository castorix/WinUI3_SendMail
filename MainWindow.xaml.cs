// Copyright (c) Microsoft Corporation and Contributors.
// Licensed under the MIT License.

using Microsoft.UI;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;

using GlobalStructures;
using static WinUI3_SendMail.MAPI;
using CDO;
//using static CDO.CDOTools;
using System.Security.Cryptography;
using System.Runtime.ConstrainedExecution;
using System.Runtime.Versioning;
//using ADODB;
using System.Data;
using System.Net.Mail;
using System.Net.Mime;
using Windows.ApplicationModel.DataTransfer;
using System.Security;
using System.Runtime.CompilerServices;
using static WinUI3_SendMail.MainWindow;
using WinRT;
using static WinUI3_SendMail.Outlook;
using System.Security.Claims;
using System.Diagnostics;

using ADODB;
using Microsoft.UI.Xaml.Shapes;
using System.Collections;
using System.Transactions;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Net.Security;
using Windows.ApplicationModel.Email;
using Windows.Web;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace WinUI3_SendMail
{
    /// <summary>
    /// An empty window that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainWindow : Window
    {
        [ComImport]
        [Guid("00000122-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IDropTarget
        {
            HRESULT DragEnter(
                [In] System.Runtime.InteropServices.ComTypes.IDataObject pDataObj,
                [In] int grfKeyState,
                [In] System.Drawing.Point pt,
                [In, Out] ref int pdwEffect);

            HRESULT DragOver(
                [In] int grfKeyState,
                [In] System.Drawing.Point pt,
                [In, Out] ref int pdwEffect);

            HRESULT DragLeave();

            HRESULT Drop(
                [In] System.Runtime.InteropServices.ComTypes.IDataObject pDataObj,
                [In] int grfKeyState,
                [In] System.Drawing.Point pt,
                [In, Out] ref int pdwEffect);
        }

        public const int DROPEFFECT_NONE = (0);

        [DllImport("Shell32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "#740")]
        public static extern HRESULT SHCreateFileDataObject(IntPtr pidlFolder, uint cidl, IntPtr[] apidl, System.Runtime.InteropServices.ComTypes.IDataObject pdtInner, out System.Runtime.InteropServices.ComTypes.IDataObject ppdtobj);

        [DllImport("Shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern HRESULT SHILCreateFromPath([MarshalAs(UnmanagedType.LPWStr)] string pszPath, out IntPtr ppIdl, ref uint rgflnOut);

        [DllImport("Shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr ILFindLastID(IntPtr pidl);

        [DllImport("Shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr ILClone(IntPtr pidl);

        [DllImport("Shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern Boolean ILRemoveLastID(IntPtr pidl);

        [DllImport("Shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern void ILFree(IntPtr pidl);

        Guid CLSID_MapiMail = new Guid("9E56BE60-C50F-11CF-9A2C-00A0C90A90CE");

        [DllImport("Kernel32.dll", SetLastError = true, CharSet = CharSet.Ansi)]
        public static extern IntPtr GetProcAddress(IntPtr hModule, string lpProcName);

        [DllImport("Kernel32.dll", SetLastError = true, CharSet = CharSet.Ansi, EntryPoint = "GetProcAddress")]
        public static extern IntPtr GetProcAddressByOrdinal(IntPtr hModule, int lpProcName);

        [DllImport("Kernel32.dll", SetLastError = true, CharSet = CharSet.Auto, BestFitMapping = false, ThrowOnUnmappableChar = true)]
        public static extern IntPtr LoadLibrary(string lpLibFileName);

        [DllImport("Oleaut32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr SysAllocStringLen(string src, int len);  // BSTR

        [DllImport("Oleaut32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr SysAllocString([In, MarshalAs(UnmanagedType.LPWStr)] string s);

        [DllImport("Oleaut32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern void SysFreeString(IntPtr pbstr);

        [DllImport("Oleaut32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern HRESULT GetActiveObject(ref Guid rclsid, IntPtr pvReserved, out IntPtr ppunk);

        [DllImport("Ole32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern HRESULT CLSIDFromProgID(string lpszProgID, out Guid lpclsid);

        [DllImport("Shell32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern HRESULT SHGetKnownFolderPath([In, MarshalAs(UnmanagedType.LPStruct)] Guid rfid, int dwFlags, IntPtr hToken, out IntPtr ppszPath);

        public enum KNOWN_FOLDER_FLAG : uint
        {
            KF_FLAG_DEFAULT = 0x00000000,
            KF_FLAG_FORCE_APP_DATA_REDIRECTION = 0x00080000,
            KF_FLAG_RETURN_FILTER_REDIRECTION_TARGET = 0x00040000,
            KF_FLAG_FORCE_PACKAGE_REDIRECTION = 0x00020000,
            KF_FLAG_NO_PACKAGE_REDIRECTION = 0x00010000,
            KF_FLAG_FORCE_APPCONTAINER_REDIRECTION = 0x00020000,
            KF_FLAG_NO_APPCONTAINER_REDIRECTION = 0x00010000,
            KF_FLAG_CREATE = 0x00008000,
            KF_FLAG_DONT_VERIFY = 0x00004000,
            KF_FLAG_DONT_UNEXPAND = 0x00002000,
            KF_FLAG_NO_ALIAS = 0x00001000,
            KF_FLAG_INIT = 0x00000800,
            KF_FLAG_DEFAULT_PATH = 0x00000400,
            KF_FLAG_NOT_PARENT_RELATIVE = 0x00000200,
            KF_FLAG_SIMPLE_IDLIST = 0x00000100,
            KF_FLAG_ALIAS_ONLY = 0x80000000,
        }

        public static readonly Guid FOLDERID_ProgramFiles = new Guid("905e63b6-c1bf-494e-b29c-65b732d3d21a");
        //C:\Program Files(x86)\Internet Explorer %ProgramFiles% (%SystemDrive%\Program Files)

        // https://github.com/andrewleader/WindowsAppSDKGallery/blob/main/WindowsAppSDKGallery/SamplePages/ShareSamples/DataTransferManagerPage.xaml.cs#L39
        [ComImport]
        [Guid("3A3DCD6C-3EAB-43DC-BCDE-45671CE800C8")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IDataTransferManagerInterop
        {
            HRESULT GetForWindow(IntPtr appWindow, ref Guid riid, out IntPtr dataTransferManager);
            HRESULT ShowShareUIForWindow(IntPtr appWindow);
        }
     
        public static readonly Guid Guid_DTM = new Guid(0xa5caee9b, 0x8708, 0x49d1, 0x8d, 0x36, 0x67, 0xd2, 0x5a, 0x8d, 0xa0, 0x0c);
        static IDataTransferManagerInterop DataTransferManagerInterop => DataTransferManager.As<IDataTransferManagerInterop>();

        System.Collections.ObjectModel.ObservableCollection<string> files = new System.Collections.ObjectModel.ObservableCollection<string>();
        List<Windows.Storage.IStorageItem> filesToShare = new List<Windows.Storage.IStorageItem>();
        private IntPtr hWnd = IntPtr.Zero;
        private Microsoft.UI.Windowing.AppWindow _apw;
        private Microsoft.UI.Xaml.DispatcherTimer dTimer;

        public MainWindow()
        {
            this.InitializeComponent();

            hWnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
            Microsoft.UI.WindowId myWndId = Microsoft.UI.Win32Interop.GetWindowIdFromWindow(hWnd);
            _apw = Microsoft.UI.Windowing.AppWindow.GetFromWindowId(myWndId);
            _apw.Resize(new Windows.Graphics.SizeInt32(1000, 790));
            _apw.Move(new Windows.Graphics.PointInt32(400, 150));
            Application.Current.Resources["ButtonBackgroundPressed"] = new SolidColorBrush(Microsoft.UI.Colors.LightSteelBlue);
            Application.Current.Resources["ButtonBackgroundPointerOver"] = new SolidColorBrush(Microsoft.UI.Colors.RoyalBlue);

            string sDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string sPath = sDirectory + @"Assets\Test.docx";
            files.Add(sPath);
            sPath = sDirectory + @"Assets\Butterfly.png";
            files.Add(sPath);
 
            AffFilesToShare(filesToShare);            

            string sKey = @"SOFTWARE\Microsoft\Office";
            bool bFound = false;
            string sEmail = FindValue(RegistryHive.CurrentUser, sKey, "EmailAddress", ref bFound);
            if (sEmail != null)
            {
                tbSender.Text = sEmail;
                tbRecipient.Text = sEmail;
            }
            bFound = false;
            string sSMTPServer = FindValue(RegistryHive.CurrentUser, sKey, "SMTP Server", ref bFound);
            tbSMTPServer.Text = sSMTPServer;
            bFound = false;
            string sSMTPPort = FindValue(RegistryHive.CurrentUser, sKey, "SMTP Port", ref bFound);
            tbSMTPServerPort.Text = sSMTPPort;
            tbSMTPUser.Text = tbSender.Text;

            // For bug on ContextMenu on (Rich)TextBox controls
            //https://github.com/microsoft/microsoft-ui-xaml/issues/4804
            tbSender.ContextFlyout.Opening += Menu_Opening;
            tbSender.SelectionFlyout.Opening += Menu_Opening;

            // https://github.com/microsoft/Windows-classic-samples/blob/main/Samples/ShareSource/wpf/MainWindow.xaml.cs
            IntPtr pDTM = IntPtr.Zero;          
            HRESULT hr = DataTransferManagerInterop.GetForWindow(hWnd, Guid_DTM, out pDTM);            
            DataTransferManager dtm = MarshalInterface<DataTransferManager>.FromAbi(pDTM);
            dtm.DataRequested += (s, args) =>
            {
                var deferral = args.Request.GetDeferral();
                try
                {
                    DataPackage dp = args.Request.Data;
                    if (tbSubject.Text != "")
                        dp.Properties.Title = tbSubject.Text;
                    else
                        dp.Properties.Title = "This is a test";
                    //dp.RequestedOperation = DataPackageOperation.Link;
                    dp.RequestedOperation = DataPackageOperation.None;
                    dp.SetStorageItems(filesToShare);
                }
                finally
                {
                    deferral.Complete();
                }
            }; 
        }

        private async void AffFilesToShare(List<Windows.Storage.IStorageItem>filesToShare)
        {           
            for (int i = 0; i < files.Count; i++)
            {
                Windows.Storage.StorageFile file = await Windows.Storage.StorageFile.GetFileFromPathAsync(files[i]);
                filesToShare.Add(file);
            }
        }

        public string FindValue(RegistryHive rh, string sKey, string sSearchValue, ref bool bFound)
        {
            string sReturn = null;
            using (RegistryKey rkCCU = RegistryKey.OpenBaseKey(rh, RegistryView.Registry64))  
            {               
                RegistryKey rk = rkCCU.OpenSubKey(sKey, false);
                string[] sSubKeys = rk.GetSubKeyNames();
                foreach (string sSubKey in sSubKeys)
                {
                    string sFullKey = sKey + "\\" + sSubKey;
                    using (RegistryKey rkSubKey = rkCCU.OpenSubKey(sFullKey, false))
                    {
                        string[] sValues = rkSubKey.GetValueNames();
                        foreach (string sValue in sValues)
                        {
                            if (sValue == sSearchValue)
                            {
                                RegistryValueKind rvk = rkSubKey.GetValueKind(sValue);
                                if (rvk == RegistryValueKind.String)
                                    sReturn = (string)rkSubKey.GetValue(sValue, 0);
                                else if (rvk == RegistryValueKind.DWord)
                                    sReturn = (rkSubKey.GetValue(sValue, 0)).ToString();
                                // + Binary...
                                bFound = true;
                                return sReturn;
                            }
                            //System.Diagnostics.Debug.WriteLine(sValue);
                        }
                    }
                    if (!bFound)
                        sReturn = FindValue(rh, sFullKey, sSearchValue, ref bFound);
                    else
                        break;
                    //System.Diagnostics.Debug.WriteLine(sFullKey);
                }
            }
            return sReturn;
        }

        // https://www.unicode.org/emoji/charts/full-emoji-list.html
        private async void btnSendMail_Click(object sender, RoutedEventArgs e)
        {
            if (g_nSendMail == SENDMAIL.MAPI)
            {
                if (g_nMAPISendMail == MAPISENDMAIL.DATAOBJECT)
                {
                    HRESULT hr = HRESULT.E_FAIL;
                    uint rgflnOut = 0;
                    System.Runtime.InteropServices.ComTypes.IDataObject pDataObject;
                    IntPtr pidlParent = IntPtr.Zero, pidlFull = IntPtr.Zero, pidlItem = IntPtr.Zero;
                    var aPidl = new IntPtr[255];
                    int nNbFiles = 0;
                    for (int i = 0; i < files.Count; i++)
                    {
                        hr = SHILCreateFromPath(files[i], out pidlFull, ref rgflnOut);
                        if (hr == HRESULT.S_OK)
                        {
                            pidlItem = ILFindLastID(pidlFull);
                            aPidl[i] = ILClone(pidlItem);
                            ILRemoveLastID(pidlFull);
                            pidlParent = ILClone(pidlFull);
                            ILFree(pidlFull);
                            nNbFiles++;
                        }
                    }

                    hr = SHCreateFileDataObject(pidlParent, (uint)nNbFiles, aPidl, null, out pDataObject);
                    if (hr == HRESULT.S_OK)
                    {
                        Type DropTargetType = Type.GetTypeFromCLSID(CLSID_MapiMail, true);
                        object DropTarget = Activator.CreateInstance(DropTargetType);
                        IDropTarget pDropTarget = (IDropTarget)DropTarget;
                        int pdwEffect = DROPEFFECT_NONE;
                        System.Drawing.Point pt = new System.Drawing.Point(0, 0);
                        hr = pDropTarget.Drop(pDataObject, 0, pt, pdwEffect);
                        Marshal.ReleaseComObject(pDataObject);
                        Marshal.ReleaseComObject(pDropTarget);
                    }

                    if (pidlParent != IntPtr.Zero)
                        ILFree(pidlParent);
                    for (int i = 0; i < nNbFiles; i++)
                    {
                        if (aPidl[i] != IntPtr.Zero)
                            ILFree(aPidl[i]);
                    }
                    StartTimer(60 * 3 * 1000, 60 * 3 * 1000);
                    //Console.Beep(5000, 10);
                }
                else if (g_nMAPISendMail == MAPISENDMAIL.API)
                {
                    // Old HTTP MAPISendMail : Does not work anymore...
                    if (1 == 0)
                    {
                        string sFrom = tbSender.Text;
                        string sTo = tbRecipient.Text;
                        string sSubject = tbSubject.Text;
                        string sText = string.Empty;
                        rebText.Document.GetText(Microsoft.UI.Text.TextGetOptions.AdjustCrlf, out sText);

                        MapiRecipDesc mrdFromAnsi = new MapiRecipDesc();
                        mrdFromAnsi.lpszName = sFrom;
                        mrdFromAnsi.ulRecipClass = MAPI_ORIG;
                        MapiRecipDesc mrdToAnsi = new MapiRecipDesc();
                        mrdToAnsi.lpszName = sTo;
                        mrdToAnsi.ulRecipClass = MAPI_TO;
                        MapiMessage mmAnsi = new MapiMessage();
                        mmAnsi.lpszSubject = sSubject;
                        mmAnsi.lpszNoteText = sText;

                        IntPtr pFromAnsi = Marshal.AllocHGlobal(Marshal.SizeOf(mrdFromAnsi));
                        Marshal.StructureToPtr(mrdFromAnsi, pFromAnsi, false);
                        mmAnsi.lpOriginator = pFromAnsi;
                        IntPtr pToAnsi = Marshal.AllocHGlobal(Marshal.SizeOf(mrdToAnsi));
                        Marshal.StructureToPtr(mrdToAnsi, pToAnsi, false);
                        mmAnsi.lpRecips = pToAnsi;
                        mmAnsi.nRecipCount = 1;

                        List<MapiFileDesc> listAnsi = new List<MapiFileDesc>();
                        for (int i = 0; i < files.Count; i++)
                        {
                            MapiFileDesc mfd = new MapiFileDesc(0, 0, 0xFFFFFFFF, files[i], null, IntPtr.Zero);
                            listAnsi.Add(mfd);
                        }
                        MapiFileDesc[] aMFDAnsi = listAnsi.ToArray();

                        int nStructSizeAnsi = Marshal.SizeOf(typeof(MapiFileDesc));
                        IntPtr pArrayAnsi = Marshal.AllocHGlobal(aMFDAnsi.Length * nStructSizeAnsi);
                        IntPtr ptrAnsi = pArrayAnsi;
                        for (int i = 0; i < aMFDAnsi.Length; i++)
                        {
                            Marshal.StructureToPtr(aMFDAnsi[i], ptrAnsi, false);
                            ptrAnsi += nStructSizeAnsi;
                        }
                        mmAnsi.lpFiles = pArrayAnsi;
                        mmAnsi.nFileCount = (uint)aMFDAnsi.Length;

                        bool bError = false;
                        try
                        {
                            IntPtr pMessageAnsi = Marshal.AllocHGlobal(Marshal.SizeOf(mmAnsi));
                            Marshal.StructureToPtr(mmAnsi, pMessageAnsi, false);
                            IntPtr pPath = IntPtr.Zero;
                            HRESULT hr = SHGetKnownFolderPath(FOLDERID_ProgramFiles, 0, IntPtr.Zero, out pPath);
                            string sPath = Marshal.PtrToStringUni(pPath);
                            string sDLL = sPath + "\\Internet Explorer\\hmmapi.dll";
                            IntPtr hMapiDLL = LoadLibrary(sDLL);
                            IntPtr pMAPISendMail = GetProcAddress(hMapiDLL, "MAPISendMail");
                            MAPISendMailDelegate pMAPISendMailDelegate = (MAPISendMailDelegate)Marshal.GetDelegateForFunctionPointer(pMAPISendMail, typeof(MAPISendMailDelegate));
                            MAPI_ERROR nRet = pMAPISendMailDelegate(IntPtr.Zero, UIntPtr.Zero, pMessageAnsi, 0, 0);
                            Marshal.FreeHGlobal(pArrayAnsi);
                            Marshal.FreeHGlobal(pFromAnsi);
                            Marshal.FreeHGlobal(pToAnsi);
                            Marshal.FreeHGlobal(pMessageAnsi);
                            if (nRet != MAPI_ERROR.SUCCESS_SUCCESS)
                            {
                                if (nRet != MAPI_ERROR.MAPI_USER_ABORT)
                                {
                                    string sError = "Error : " + nRet.ToString();
                                    Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog(sError, "Error");
                                    WinRT.Interop.InitializeWithWindow.Initialize(md, hWnd);
                                    _ = await md.ShowAsync();
                                }
                                else
                                {
                                    string sError = "Mail cancelled by user";
                                    Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog(sError, "Information");
                                    WinRT.Interop.InitializeWithWindow.Initialize(md, hWnd);
                                    _ = await md.ShowAsync();
                                }
                            }
                            else
                            {
                                StartTimer(10 * 1000, 10 * 1000);
                                string sError = "Mail sent !";
                                Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog(sError, "Information");
                                WinRT.Interop.InitializeWithWindow.Initialize(md, hWnd);
                                _ = await md.ShowAsync();
                            }
                        }
                        catch (Exception ex)
                        {
                            bError = true;
                            string sError = ex.Message + "\r\n" + "HRESULT = 0x" + string.Format("{0:X}", ex.HResult);
                            Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog(sError, "Error");
                            WinRT.Interop.InitializeWithWindow.Initialize(md, hWnd);
                            _ = await md.ShowAsync();
                        }
                        finally
                        {
                            if (!bError)
                            {
                                string sError = "Mail sent !";
                                Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog(sError, "Information");
                                WinRT.Interop.InitializeWithWindow.Initialize(md, hWnd);
                                _ = await md.ShowAsync();
                            }                           
                        } 
                    }
                    else
                    {
                        // No dialog : C:\WINDOWS\SysWOW64\fixmapi.exe -Embedding
                        string sFrom = tbSender.Text;
                        string sTo = tbRecipient.Text;
                        string sSubject = tbSubject.Text;
                        string sText = string.Empty;
                        rebText.Document.GetText(Microsoft.UI.Text.TextGetOptions.AdjustCrlf, out sText);

                        IntPtr hMapiDLL = LoadDefaultMailProvider();
                        // Unicode
                        IntPtr pMAPISendMail = GetProcAddress(hMapiDLL, "MAPISendMailW");
                        MAPISendMailDelegate pMAPISendMailWDelegate = (MAPISendMailDelegate)Marshal.GetDelegateForFunctionPointer(pMAPISendMail, typeof(MAPISendMailDelegate));

                        MapiRecipDescW mrdFrom = new MapiRecipDescW();
                        mrdFrom.lpszName = sFrom;
                        mrdFrom.ulRecipClass = MAPI_ORIG;
                        MapiRecipDescW mrdTo = new MapiRecipDescW();
                        mrdTo.lpszName = sTo;
                        mrdTo.ulRecipClass = MAPI_TO;
                        MapiMessageW mm = new MapiMessageW();
                        mm.lpszSubject = sSubject;
                        mm.lpszNoteText = sText;

                        IntPtr pFrom = Marshal.AllocHGlobal(Marshal.SizeOf(mrdFrom));
                        Marshal.StructureToPtr(mrdFrom, pFrom, false);
                        mm.lpOriginator = pFrom;
                        IntPtr pTo = Marshal.AllocHGlobal(Marshal.SizeOf(mrdTo));
                        Marshal.StructureToPtr(mrdTo, pTo, false);
                        mm.lpRecips = pTo;
                        mm.nRecipCount = 1;

                        List<MapiFileDescW> list = new List<MapiFileDescW>();
                        for (int i = 0; i < files.Count; i++)
                        {
                            MapiFileDescW mfd = new MapiFileDescW(0, 0, 0xFFFFFFFF, files[i], null, IntPtr.Zero);
                            list.Add(mfd);
                        }
                        MapiFileDescW[] aMFD = list.ToArray();

                        int nStructSize = Marshal.SizeOf(typeof(MapiFileDescW));
                        IntPtr pArray = Marshal.AllocHGlobal(aMFD.Length * nStructSize);
                        IntPtr ptr = pArray;
                        for (int i = 0; i < aMFD.Length; i++)
                        {
                            Marshal.StructureToPtr(aMFD[i], ptr, false);
                            ptr += nStructSize;
                        }
                        mm.lpFiles = pArray;
                        mm.nFileCount = (uint)aMFD.Length;

                        IntPtr pMessage = Marshal.AllocHGlobal(Marshal.SizeOf(mm));
                        Marshal.StructureToPtr(mm, pMessage, false);

                        // Microsoft Outlook OK, but needs to be opened to synchronize
                        // SeaMonkey nRet = 0x800706be with MAPISendMailW
                        //MAPI_ERROR nRet = pMAPISendMailWDelegate(IntPtr.Zero, UIntPtr.Zero, pMessage, MAPI_DIALOG | MAPI_NEW_SESSION | MAPI_LOGON_UI, 0);
                        int nFlags = MAPI_LOGON_UI;
                        if ((bool)cbDialog.IsChecked)
                            nFlags += MAPI_DIALOG;
                        MAPI_ERROR nRet = pMAPISendMailWDelegate(IntPtr.Zero, UIntPtr.Zero, pMessage, nFlags, 0);
                        Marshal.FreeHGlobal(pArray);
                        Marshal.FreeHGlobal(pFrom);
                        Marshal.FreeHGlobal(pTo);
                        Marshal.FreeHGlobal(pMessage);
                        if (nRet != MAPI_ERROR.SUCCESS_SUCCESS)
                        {
                            if (nRet != MAPI_ERROR.MAPI_USER_ABORT)
                            {
                                // Ansi
                                pMAPISendMail = GetProcAddress(hMapiDLL, "MAPISendMail");
                                pMAPISendMailWDelegate = (MAPISendMailDelegate)Marshal.GetDelegateForFunctionPointer(pMAPISendMail, typeof(MAPISendMailDelegate));

                                MapiRecipDesc mrdFromAnsi = new MapiRecipDesc();
                                mrdFromAnsi.lpszName = sFrom;
                                mrdFromAnsi.ulRecipClass = MAPI_ORIG;
                                MapiRecipDesc mrdToAnsi = new MapiRecipDesc();
                                mrdToAnsi.lpszName = sTo;
                                mrdToAnsi.ulRecipClass = MAPI_TO;
                                MapiMessage mmAnsi = new MapiMessage();
                                mmAnsi.lpszSubject = sSubject;
                                mmAnsi.lpszNoteText = sText;

                                IntPtr pFromAnsi = Marshal.AllocHGlobal(Marshal.SizeOf(mrdFromAnsi));
                                Marshal.StructureToPtr(mrdFromAnsi, pFromAnsi, false);
                                mmAnsi.lpOriginator = pFromAnsi;
                                IntPtr pToAnsi = Marshal.AllocHGlobal(Marshal.SizeOf(mrdToAnsi));
                                Marshal.StructureToPtr(mrdToAnsi, pToAnsi, false);
                                mmAnsi.lpRecips = pToAnsi;
                                mmAnsi.nRecipCount = 1;

                                List<MapiFileDesc> listAnsi = new List<MapiFileDesc>();
                                for (int i = 0; i < files.Count; i++)
                                {
                                    MapiFileDesc mfd = new MapiFileDesc(0, 0, 0xFFFFFFFF, files[i], null, IntPtr.Zero);
                                    listAnsi.Add(mfd);
                                }
                                MapiFileDesc[] aMFDAnsi = listAnsi.ToArray();

                                int nStructSizeAnsi = Marshal.SizeOf(typeof(MapiFileDesc));
                                IntPtr pArrayAnsi = Marshal.AllocHGlobal(aMFDAnsi.Length * nStructSizeAnsi);
                                IntPtr ptrAnsi = pArrayAnsi;
                                for (int i = 0; i < aMFDAnsi.Length; i++)
                                {
                                    Marshal.StructureToPtr(aMFDAnsi[i], ptrAnsi, false);
                                    ptrAnsi += nStructSizeAnsi;
                                }
                                mmAnsi.lpFiles = pArrayAnsi;
                                mmAnsi.nFileCount = (uint)aMFDAnsi.Length;

                                IntPtr pMessageAnsi = Marshal.AllocHGlobal(Marshal.SizeOf(mmAnsi));
                                Marshal.StructureToPtr(mmAnsi, pMessageAnsi, false);
                                nRet = pMAPISendMailWDelegate(IntPtr.Zero, UIntPtr.Zero, pMessageAnsi, nFlags, 0);
                                Marshal.FreeHGlobal(pArrayAnsi);
                                Marshal.FreeHGlobal(pFromAnsi);
                                Marshal.FreeHGlobal(pToAnsi);
                                Marshal.FreeHGlobal(pMessageAnsi);
                                if (nRet != MAPI_ERROR.SUCCESS_SUCCESS)
                                {
                                    if (nRet != MAPI_ERROR.MAPI_USER_ABORT)
                                    {
                                        string sError = "Error : " + nRet.ToString();
                                        Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog(sError, "Error");
                                        WinRT.Interop.InitializeWithWindow.Initialize(md, hWnd);
                                        _ = await md.ShowAsync();
                                    }
                                    else
                                    {
                                        string sError = "Mail cancelled by user";
                                        Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog(sError, "Information");
                                        WinRT.Interop.InitializeWithWindow.Initialize(md, hWnd);
                                        _ = await md.ShowAsync();
                                    }
                                }
                                else
                                {
                                    StartTimer(10 * 1000, 10 * 1000);
                                    string sError = "Mail sent !";
                                    Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog(sError, "Information");
                                    WinRT.Interop.InitializeWithWindow.Initialize(md, hWnd);
                                    _ = await md.ShowAsync();
                                }
                            }
                            else
                            {
                                string sError = "Mail cancelled by user";
                                Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog(sError, "Information");
                                WinRT.Interop.InitializeWithWindow.Initialize(md, hWnd);
                                _ = await md.ShowAsync();
                            }
                        }
                        else
                        {
                            StartTimer(10 * 1000, 10 * 1000);
                            string sError = "Mail sent !";
                            Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog(sError, "Information");
                            WinRT.Interop.InitializeWithWindow.Initialize(md, hWnd);
                            _ = await md.ShowAsync();
                        }
                    }
                }
            }
            else if (g_nSendMail == SENDMAIL.CDO)
            {
                Message msg = new Message();
                msg.Sender = tbSender.Text;
                msg.To = tbRecipient.Text;
                msg.Subject = tbSubject.Text;
                string sText = string.Empty;
                rebText.Document.GetText(Microsoft.UI.Text.TextGetOptions.AdjustCrlf, out sText);
                msg.TextBody = sText;
                msg.TextBodyPart.Charset = "utf-8";

                ADODB.Fields msgFields = msg.Fields;
                ADODB.Field msgField = msgFields["urn:schemas:httpmail:priority"];
                msgField.Value = cdoPriorityValues.cdoPriorityUrgent;
                msgField = msgFields["urn:schemas:httpmail:importance"];
                msgField.Value = cdoImportanceValues.cdoHigh;
                msgFields.Update();
                Marshal.ReleaseComObject(msgFields);

                for (int i = 0; i < files.Count; i++)
                {
                    msg.AddAttachment(files[i], "", "");
                }

                var config = new Configuration();
                ADODB.Fields fields = config.Fields;
                var n = fields.Count;
                string sSchema = "http://schemas.microsoft.com/cdo/configuration/";
                ADODB.Field field = fields[sSchema + "sendusing"];
                field.Value = CdoSendUsing.cdoSendUsingPort;
                field = fields[sSchema + "smtpserver"];
                field.Value = tbSMTPServer.Text;
                field = fields[sSchema + "smtpserverport"];
                field.Value = tbSMTPServerPort.Text;
                //field.Value = 465;                

                if (tbSMTPUser.Text != "")
                {
                    field = fields[sSchema + "smtpauthenticate"];
                    field.Value = CdoProtocolsAuthentication.cdoBasic;
                    field = fields[sSchema + "sendusername"];
                    field.Value = tbSMTPUser.Text;
                    field = fields[sSchema + "sendpassword"];
                    field.Value = tbSMTPPassword.Text;
                    field = fields[sSchema + "smtpusessl"];
                    field.Value = true;
                }
                fields.Update();
                Marshal.ReleaseComObject(fields);

                msg.Configuration = config;

                bool bError = false;
                try
                {
                    msg.Send();
                }
                catch (Exception ex)
                {
                    bError = true;
                    string sError = ex.Message + "\r\n" + "HRESULT = 0x" + string.Format("{0:X}", ex.HResult);
                    Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog(sError, "Error");
                    WinRT.Interop.InitializeWithWindow.Initialize(md, hWnd);
                    _ = await md.ShowAsync();
                }
                finally
                {
                    if (!bError)
                    {
                        string sError = "Mail sent !";
                        Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog(sError, "Information");
                        WinRT.Interop.InitializeWithWindow.Initialize(md, hWnd);
                        _ = await md.ShowAsync();
                    }
                    Marshal.ReleaseComObject(msg);
                    Marshal.ReleaseComObject(config);
                }
            }
            else if (g_nSendMail == SENDMAIL.NETMAIL)
            {
                int nPort = 0;
                bool bRet = int.TryParse(tbSMTPServerPort.Text, out nPort);
                if (bRet)
                {
                    System.Net.Mail.SmtpClient smtpClient = new System.Net.Mail.SmtpClient(tbSMTPServer.Text, nPort);
                    System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage();

                    String sFrom = tbSender.Text;
                    mail.From = new System.Net.Mail.MailAddress(sFrom);
                    mail.To.Add(tbRecipient.Text);
                    mail.Subject = tbSubject.Text;
                    string sText = string.Empty;
                    rebText.Document.GetText(Microsoft.UI.Text.TextGetOptions.AdjustCrlf | Microsoft.UI.Text.TextGetOptions.UseCrlf, out sText);
                    mail.Body = sText;
                    mail.BodyEncoding = System.Text.Encoding.UTF8;
                    smtpClient.EnableSsl = true;
                    smtpClient.UseDefaultCredentials = false;
                    smtpClient.Credentials = new System.Net.NetworkCredential(tbSMTPUser.Text, tbSMTPPassword.Text);

                    for (int i = 0; i < files.Count; i++)
                    {
                        Attachment data = new Attachment(files[i], MediaTypeNames.Application.Octet);
                        mail.Attachments.Add(data);
                    }

                    bool bError = false;
                    try
                    {
                        smtpClient.Send(mail);
                    }
                    catch (Exception ex)
                    {
                        bError = true;
                        string sError = ex.Message + "\r\n" + "HRESULT = 0x" + string.Format("{0:X}", ex.HResult);
                        Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog(sError, "Error");
                        WinRT.Interop.InitializeWithWindow.Initialize(md, hWnd);
                        _ = await md.ShowAsync();
                    }
                    finally
                    {
                        if (!bError)
                        {
                            string sError = "Mail sent !";
                            Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog(sError, "Information");
                            WinRT.Interop.InitializeWithWindow.Initialize(md, hWnd);
                            _ = await md.ShowAsync();
                        }
                    }
                }
                else
                {
                    string sError = "SMTP Port is incorrect";
                    Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog(sError, "Information");
                    WinRT.Interop.InitializeWithWindow.Initialize(md, hWnd);
                    _ = await md.ShowAsync();
                }
            }
            else if (g_nSendMail == SENDMAIL.WINSOCK)
            {
                string sMessageData = null;
                int nPort = 0;              
                bool bRet = int.TryParse(tbSMTPServerPort.Text, out nPort);
                if (bRet)
                {   
                    bool bError = false;
                    try
                    {
                        //https://www.samlogic.net/articles/smtp-commands-reference.htm
                        using (var client = new TcpClient(tbSMTPServer.Text, nPort))
                        {
                            using (var stream = client.GetStream())
                            using (var streamReader = new StreamReader(stream))
                            using (var streamWriter = new StreamWriter(stream) { AutoFlush = true })
                            using (var sslStream = new System.Net.Security.SslStream(stream))
                            {
                                sMessageData = ReadMessage2(streamReader);
                                tbResponse.Text += sMessageData;
                                if (!sMessageData.StartsWith("220"))
                                {
                                    bError = true;
                                    throw new InvalidOperationException("Could not connect to SMTP server");
                                }

                                streamWriter.WriteLine(string.Format("HELO {0}\r\n", tbSMTPServer.Text));
                                sMessageData = ReadMessage2(streamReader);
                                tbResponse.Text += sMessageData;
                                if (!sMessageData.StartsWith("250"))
                                {
                                    bError = true;
                                    throw new InvalidOperationException("HELO command failed");
                                }

                                streamWriter.WriteLine("STARTTLS");
                                sMessageData = ReadMessage2(streamReader);
                                tbResponse.Text += sMessageData;
                                if (!sMessageData.StartsWith("220"))
                                {
                                    bError = true;
                                    throw new InvalidOperationException("STARTTLS command failed");
                                }

                                sslStream.AuthenticateAsClient(tbSMTPServer.Text);

                                {
                                    sslStream.Write(Encoding.UTF8.GetBytes(string.Format("EHLO {0}\r\n", tbSMTPServer.Text)));
                                    sMessageData = ReadMessage(sslStream);
                                    tbResponse.Text += sMessageData;
                                    if (!sMessageData.StartsWith("250"))
                                    {
                                        bError = true;
                                        throw new InvalidOperationException("EHLO command failed");
                                    }

                                    sslStream.Write(Encoding.UTF8.GetBytes(string.Format("AUTH LOGIN\r\n")));                                   
                                    sMessageData = ReadMessage(sslStream);
                                    tbResponse.Text += sMessageData;
                                    if (!sMessageData.StartsWith("334"))
                                    {
                                        bError = true;
                                        throw new InvalidOperationException("AUTH LOGIN command failed");
                                    }

                                    string sUserBase64 = System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(tbSMTPUser.Text));
                                    sslStream.Write(Encoding.UTF8.GetBytes(sUserBase64 + "\r\n"));   
                                    sMessageData = ReadMessage(sslStream);
                                    tbResponse.Text += sMessageData;
                                    if (!sMessageData.StartsWith("334"))
                                    {
                                        bError = true;
                                        throw new InvalidOperationException("Bad User");
                                    }

                                    string sPasswordBase64 = System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(tbSMTPPassword.Text));
                                    sslStream.Write(Encoding.UTF8.GetBytes(sPasswordBase64 + "\r\n"));
                                    sMessageData = ReadMessage(sslStream);
                                    tbResponse.Text += sMessageData;
                                    if (!sMessageData.StartsWith("235"))
                                    {
                                        bError = true;
                                        throw new InvalidOperationException("Bad User/Password");
                                    }

                                    sslStream.Write(Encoding.UTF8.GetBytes(string.Format("MAIL FROM:<{0}>\r\n", tbSender.Text)));
                                    sMessageData = ReadMessage(sslStream);
                                    tbResponse.Text += sMessageData;
                                    if (!sMessageData.StartsWith("250"))
                                    {
                                        bError = true;
                                        throw new InvalidOperationException("Bad Mail Sender");
                                    }

                                    sslStream.Write(Encoding.UTF8.GetBytes(string.Format("RCPT TO:<{0}>\r\n", tbRecipient.Text)));
                                    sMessageData = ReadMessage(sslStream);
                                    tbResponse.Text += sMessageData;
                                    if (!sMessageData.StartsWith("250"))
                                    {
                                        bError = true;
                                        throw new InvalidOperationException("Bad Mail Receiver");
                                    }

                                    sslStream.Write(Encoding.UTF8.GetBytes(string.Format("DATA\r\n")));
                                    sMessageData = ReadMessage(sslStream);
                                    tbResponse.Text += sMessageData;
                                    if (!sMessageData.StartsWith("354"))
                                    {
                                        bError = true;
                                        throw new InvalidOperationException("Bad Mail Data");
                                    }

                                    // https://codesnipets.wordpress.com/2010/04/21/smtp-email-using-sockets-system-web-mail/

                                    StringBuilder Header = new StringBuilder();
                                    Header.Append("From: " + tbSender.Text + "\r\n");
                                    string _To = tbRecipient.Text;                                  
                                    string[] Tos = tbRecipient.Text.Split(new char[] { ';' });
                                    Header.Append("To: ");
                                    for (int i = 0; i < Tos.Length; i++)
                                    {
                                        Header.Append(i > 0 ? "," : "");
                                        Header.Append(Tos[i]);
                                    }
                                    Header.Append("\r\n");                                   
                                    Header.Append("Date: ");
                                    Header.Append(DateTime.Now.ToString("ddd, d M y H:m:s z"));
                                    Header.Append("\r\n");
                                    Header.Append("Subject: " + tbSubject.Text + "\r\n");
                                    Header.Append("X-Mailer: WinUI 3\r\n");
                                    string sBodyText = string.Empty;
                                    rebText.Document.GetText(Microsoft.UI.Text.TextGetOptions.AdjustCrlf | Microsoft.UI.Text.TextGetOptions.UseCrlf, out sBodyText);                                 
                                    if (!sBodyText.EndsWith("\r\n"))
                                        sBodyText += "\r\n";

                                    //System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage();
                                    //for (int i = 0; i < files.Count; i++)
                                    //{
                                    //    Attachment data = new Attachment(files[i], MediaTypeNames.Application.Octet);
                                    //    mail.Attachments.Add(data);
                                    //}

                                    //if (mail.Attachments.Count > 0)
                                    if (files.Count > 0)
                                    {
                                        Header.Append("MIME-Version: 1.0\r\n");
                                        Header.Append("Content-Type: multipart/mixed; boundary=unique-boundary-1\r\n");
                                        Header.Append("\r\n");
                                        Header.Append("This is a multi-part message in MIME format.\r\n");
                                        StringBuilder sb = new StringBuilder();
                                        sb.Append("--unique-boundary-1\r\n");
                                        sb.Append("Content-Type: text/plain\r\n");
                                        sb.Append("Content-Transfer-Encoding: 7Bit\r\n");
                                        sb.Append("\r\n");
                                        sb.Append(sBodyText + "\r\n");
                                        sb.Append("\r\n");

                                        //foreach (object o in mail.Attachments)
                                        for (int n = 0; n < files.Count; n++)
                                        {
                                            //var a = o as Attachment;
                                            byte[] binaryData;
                                            //if (a != null)
                                            {
                                                //FileInfo f = new FileInfo(a.Filename);
                                                FileInfo f = new FileInfo(files[n]);
                                                sb.Append("--unique-boundary-1\r\n");
                                                sb.Append("Content-Type: application/octet-stream; file=" + f.Name + "\r\n");
                                                sb.Append("Content-Transfer-Encoding: base64\r\n");
                                                sb.Append("Content-Disposition: attachment; filename=" + f.Name + "\r\n");
                                                sb.Append("\r\n");
                                                FileStream fs = f.OpenRead();
                                                binaryData = new Byte[fs.Length];
                                                long bytesRead = fs.Read(binaryData, 0, (int)fs.Length);
                                                fs.Close();
                                                string base64String = System.Convert.ToBase64String(binaryData, 0, binaryData.Length);

                                                for (int i = 0; i < base64String.Length;)
                                                {
                                                    int nextchunk = 100;
                                                    if (base64String.Length - (i + nextchunk) < 0)
                                                        nextchunk = base64String.Length - i;
                                                    sb.Append(base64String.Substring(i, nextchunk));
                                                    sb.Append("\r\n");
                                                    i += nextchunk;
                                                }
                                                sb.Append("\r\n");
                                            }
                                        }
                                        sBodyText = sb.ToString();
                                    }

                                    Header.Append("\r\n");
                                    Header.Append(sBodyText);
                                    Header.Append(".\r\n");
                                    Header.Append("\r\n");
                                    Header.Append("\r\n");                                    

                                    sslStream.Write(Encoding.UTF8.GetBytes(Header.ToString()));
                                    sMessageData = ReadMessage(sslStream);
                                    tbResponse.Text += sMessageData;
                                    if (!sMessageData.StartsWith("250"))
                                    {
                                        bError = true;
                                        throw new InvalidOperationException("Bad Mail Data");
                                    }

                                    //sslStream.Write(Encoding.UTF8.GetBytes(string.Format("QUIT\r\n")));
                                    //sMessageData = ReadMessage(sslStream);
                                    //tbResponse.Text += sMessageData;
                                    //if (!sMessageData.StartsWith("221"))
                                    //{
                                    //    bError = true;
                                    //    throw new InvalidOperationException("Could not Quit");
                                    //}
                                }
                            }
                        } 

                        //byte[] nReceivevBytes = new byte[1024];
                        //int nTotalBytesReceived = 0;
                        //string sMessage = null;
                        //IPHostEntry IPhst = Dns.GetHostEntry(tbSMTPServer.Text);
                        //IPEndPoint endPt = new IPEndPoint(IPhst.AddressList[0], nPort);
                        //s = new Socket(endPt.AddressFamily, SocketType.Stream, ProtocolType.Tcp);
                        //s.Connect(endPt);

                        //nTotalBytesReceived = s.Receive(nReceivevBytes);
                        //tbResponse.Text += Encoding.ASCII.GetString(nReceivevBytes, 0, nTotalBytesReceived);

                        //sMessage = string.Format("HELO {0}\r\n", Dns.GetHostName());
                        //s.Send(Encoding.ASCII.GetBytes(sMessage));

                        //nTotalBytesReceived = s.Receive(nReceivevBytes);
                        //tbResponse.Text += Encoding.ASCII.GetString(nReceivevBytes, 0, nTotalBytesReceived);

                        //sMessage = string.Format("STARTTLS\r\n");
                        //s.Send(Encoding.ASCII.GetBytes(sMessage));

                        //nTotalBytesReceived = s.Receive(nReceivevBytes);
                        //tbResponse.Text += Encoding.ASCII.GetString(nReceivevBytes, 0, nTotalBytesReceived);

                        //sMessage = string.Format("AUTH LOGIN {0}\r\n", "Test");
                        //s.Send(Encoding.ASCII.GetBytes(sMessage));

                        //nTotalBytesReceived = s.Receive(nReceivevBytes);
                        //tbResponse.Text += Encoding.ASCII.GetString(nReceivevBytes, 0, nTotalBytesReceived);

                        //sMessage = string.Format("{0}\r\n", tbSMTPUser.Text);
                        //s.Send(Encoding.ASCII.GetBytes(sMessage));

                        //nTotalBytesReceived = s.Receive(nReceivevBytes);
                        //tbResponse.Text += Encoding.ASCII.GetString(nReceivevBytes, 0, nTotalBytesReceived);

                        //sMessage = string.Format("MAIL FROM: {0}\r\n",tbSender.Text);
                        //s.Send(Encoding.ASCII.GetBytes(sMessage));

                        //nTotalBytesReceived = s.Receive(nReceivevBytes);
                        //tbResponse.Text += Encoding.ASCII.GetString(nReceivevBytes, 0, nTotalBytesReceived);

                        //sMessage = string.Format("RCPT TO: {0}\r\n", tbRecipient.Text);
                        //s.Send(Encoding.ASCII.GetBytes(sMessage));

                        //nTotalBytesReceived = s.Receive(nReceivevBytes);
                        //tbResponse.Text += Encoding.ASCII.GetString(nReceivevBytes, 0, nTotalBytesReceived); 

                    }
                    catch (Exception ex)
                    {
                        bError = true;
                        string sError = ex.Message + "\r\n" + "HRESULT = 0x" + string.Format("{0:X}", ex.HResult);
                        Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog(sError, "Error");
                        WinRT.Interop.InitializeWithWindow.Initialize(md, hWnd);
                        _ = await md.ShowAsync();
                    }
                    finally
                    {
                        //s.Close();
                        if (!bError)
                        {                          
                            string sError = "Mail sent !";
                            Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog(sError, "Information");
                            WinRT.Interop.InitializeWithWindow.Initialize(md, hWnd);
                            _ = await md.ShowAsync();
                        }
                    }
                }
                else
                {
                    string sError = "SMTP Port is incorrect";
                    Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog(sError, "Information");
                    WinRT.Interop.InitializeWithWindow.Initialize(md, hWnd);
                    _ = await md.ShowAsync();
                }               
            }
        }

        //https://learn.microsoft.com/fr-fr/dotnet/api/system.net.security.sslstream?view=net-7.0
        static string ReadMessage(System.Net.Security.SslStream sslStream)
        {
            // Read the  message sent by the server.
            // The end of the message is signaled using the
            // "<EOF>" marker.
            byte[] buffer = new byte[2048];
            StringBuilder sMessageData = new StringBuilder();
            int bytes = -1;
            do
            {
                bytes = sslStream.Read(buffer, 0, buffer.Length);

                // Use Decoder class to convert from bytes to UTF8
                // in case a character spans two buffers.
                Decoder decoder = Encoding.UTF8.GetDecoder();
                char[] chars = new char[decoder.GetCharCount(buffer, 0, bytes)];
                decoder.GetChars(buffer, 0, bytes, chars, 0);
                sMessageData.Append(chars);
                // Check for EOF.
                if (sMessageData.ToString().IndexOf("\0") != -1)
                {
                    break;
                }
            } while (bytes != 0);
            return sMessageData.ToString();
        }

        static string ReadMessage2(StreamReader sslStream)
        {
            char[] chars = new char[2048];
            StringBuilder sMessageData = new StringBuilder();
            int nChars = -1;
            do
            {                
                nChars = sslStream.Read(chars, 0, chars.Length);  
                sMessageData.Append(chars);
                if (sMessageData.ToString().IndexOf("\0") != -1)
                {
                    break;
                }
            } while (nChars != 0);
            return sMessageData.ToString();
        }    

        private void tbResponse_TextChanged(object sender, TextChangedEventArgs e)
        {
            var grid = (Grid)VisualTreeHelper.GetChild((TextBox)sender, 0);
            for (var i = 0; i <= VisualTreeHelper.GetChildrenCount(grid) - 1; i++)
            {
                object obj = VisualTreeHelper.GetChild(grid, i);
                if (!(obj is ScrollViewer)) continue;
                ((ScrollViewer)obj).ChangeView(0.0f, ((ScrollViewer)obj).ExtentHeight, 1.0f);
                break;
            }
        }

        private void tsMAPI_Toggled(object sender, RoutedEventArgs e)
        {
            ToggleSwitch ts = sender as ToggleSwitch;
            g_nMAPISendMail = (ts.IsOn?MAPISENDMAIL.API:MAPISENDMAIL.DATAOBJECT);
            cbDialog.Visibility = (g_nMAPISendMail == MAPISENDMAIL.API)?Visibility.Visible:Visibility.Collapsed;

            tbSender.Visibility = (g_nMAPISendMail == MAPISENDMAIL.API) ? Visibility.Visible : Visibility.Collapsed;
            tbRecipient.Visibility = (g_nMAPISendMail == MAPISENDMAIL.API) ? Visibility.Visible : Visibility.Collapsed;
            tbSubject.Visibility = (g_nMAPISendMail == MAPISENDMAIL.API) ? Visibility.Visible : Visibility.Collapsed;
            rebText.Visibility = (g_nMAPISendMail == MAPISENDMAIL.API) ? Visibility.Visible : Visibility.Collapsed;
        }

        private enum SENDMAIL : int
        {
            MAPI = 0,
            CDO = 1,
            NETMAIL = 2,
            WINSOCK = 3
        }
        private enum MAPISENDMAIL : int
        {
            DATAOBJECT = 0,
            API = 1
        }

        SENDMAIL g_nSendMail = SENDMAIL.MAPI;
        MAPISENDMAIL g_nMAPISendMail = MAPISENDMAIL.DATAOBJECT;

        private void rbSendMail_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is RadioButtons rb)
            {
                if (rb.SelectedItem != null)
                {
                    string sName = ((Microsoft.UI.Xaml.FrameworkElement)rb.SelectedItem).Name;
                    switch (sName)
                    {
                        case "rbMAPI":
                            g_nSendMail = SENDMAIL.MAPI;
                            cbDialog.Visibility = (g_nMAPISendMail == MAPISENDMAIL.API) ? Visibility.Visible : Visibility.Collapsed;
                            cbOutlookRefresh.Visibility = Visibility.Visible;

                            tbSender.Visibility = (g_nMAPISendMail == MAPISENDMAIL.API) ? Visibility.Visible : Visibility.Collapsed;
                            tbRecipient.Visibility = (g_nMAPISendMail == MAPISENDMAIL.API) ? Visibility.Visible : Visibility.Collapsed;
                            tbSubject.Visibility = (g_nMAPISendMail == MAPISENDMAIL.API) ? Visibility.Visible : Visibility.Collapsed;
                            rebText.Visibility = (g_nMAPISendMail == MAPISENDMAIL.API) ? Visibility.Visible : Visibility.Collapsed;

                            tbSMTPServer.Visibility = Visibility.Collapsed;
                            tbSMTPServerPort.Visibility = Visibility.Collapsed;
                            tbSMTPUser.Visibility = Visibility.Collapsed;
                            tbSMTPPassword.Visibility = Visibility.Collapsed;

                            tbResponse.Visibility = Visibility.Collapsed;
                            break;
                        case "rbCDO":
                            g_nSendMail = SENDMAIL.CDO;
                            cbDialog.Visibility = Visibility.Collapsed;
                            cbOutlookRefresh.Visibility = Visibility.Collapsed;

                            tbSender.Visibility = Visibility.Visible;
                            tbRecipient.Visibility = Visibility.Visible;
                            tbSubject.Visibility = Visibility.Visible;
                            rebText.Visibility = Visibility.Visible;

                            tbSMTPServer.Visibility = Visibility.Visible;
                            tbSMTPServerPort.Visibility = Visibility.Visible;
                            tbSMTPUser.Visibility = Visibility.Visible;
                            tbSMTPPassword.Visibility = Visibility.Visible;

                            tbResponse.Visibility = Visibility.Collapsed;
                            break;
                        case "rbNetMail":
                            g_nSendMail = SENDMAIL.NETMAIL;
                            cbDialog.Visibility = Visibility.Collapsed;
                            cbOutlookRefresh.Visibility = Visibility.Collapsed;

                            tbSender.Visibility = Visibility.Visible;
                            tbRecipient.Visibility = Visibility.Visible;
                            tbSubject.Visibility = Visibility.Visible;
                            rebText.Visibility = Visibility.Visible;

                            tbSMTPServer.Visibility = Visibility.Visible;
                            tbSMTPServerPort.Visibility = Visibility.Visible;
                            tbSMTPUser.Visibility = Visibility.Visible;
                            tbSMTPPassword.Visibility = Visibility.Visible;

                            tbResponse.Visibility = Visibility.Collapsed;
                            break;
                        case "rbWinsock":
                            g_nSendMail = SENDMAIL.WINSOCK;
                            cbDialog.Visibility = Visibility.Collapsed;
                            cbOutlookRefresh.Visibility = Visibility.Collapsed;

                            tbSender.Visibility = Visibility.Visible;
                            tbRecipient.Visibility = Visibility.Visible;
                            tbSubject.Visibility = Visibility.Visible;
                            rebText.Visibility = Visibility.Visible;

                            tbSMTPServer.Visibility = Visibility.Visible;
                            tbSMTPServerPort.Visibility = Visibility.Visible;
                            tbSMTPUser.Visibility = Visibility.Visible;
                            tbSMTPPassword.Visibility = Visibility.Visible;

                            tbResponse.Visibility = Visibility.Visible;
                            break;
                    }
                }
            }
        }

        // Simplified from "MapiUnicodeHelp.h"
        private IntPtr LoadDefaultMailProvider()
        {
            IntPtr hDLL = IntPtr.Zero;
            string szDefaultMail = null;
            string szDLLPath = null;
            using (RegistryKey rkCCU = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64))
            {
                using (RegistryKey rk = rkCCU.OpenSubKey("Software\\Clients\\Mail", false))
                {
                    if (rk != null)
                    {
                        szDefaultMail = (string)rk.GetValue("", 0);
                        if (szDefaultMail != null)
                        {
                            string szDefaultMailKey = "Software\\Clients\\Mail\\";
                            szDefaultMailKey += szDefaultMail;
                            using (RegistryKey rkDefaultMail = rkCCU.OpenSubKey(szDefaultMailKey, false))
                            {
                                if (rkDefaultMail == null)
                                {
                                    using (RegistryKey rkLocal2 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
                                    {
                                        using (RegistryKey rkDefaultMail2 = rkLocal2.OpenSubKey(szDefaultMailKey, false))
                                        {
                                            if (rkDefaultMail2 != null)
                                                szDLLPath = (string)rkDefaultMail2.GetValue("DLLPath", 0);
                                        }
                                    }
                                }
                                else
                                    szDLLPath = (string)rkDefaultMail.GetValue("DLLPath", 0);
                            }
                        }
                    }
                    else
                    {
                        using (RegistryKey rkLM = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
                        {
                            using (RegistryKey rk2 = rkLM.OpenSubKey("Software\\Clients\\Mail", false))
                            {
                                if (rk2 != null)
                                {
                                    szDefaultMail = (string)rk2.GetValue("", 0);
                                    if (szDefaultMail != null)
                                    {
                                        string szDefaultMailKey = "Software\\Clients\\Mail\\";
                                        szDefaultMailKey += szDefaultMail;
                                        using (RegistryKey rkDefaultMail = rkLM.OpenSubKey(szDefaultMailKey, false))
                                        {
                                            if (rkDefaultMail != null)
                                            {
                                                szDLLPath = (string)rkDefaultMail.GetValue("DLLPath", 0);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            if (szDLLPath != null)
            {
                hDLL = LoadLibrary(szDLLPath);
            }
            else
                hDLL = LoadLibrary("MAPI32.DLL");
            return hDLL;
        }

        private System.Threading.Tasks.Task RefreshOutlook()
        {
            _Application pOutlook = null;
            _NameSpace ns = null;
            bool bOutlookRunning = false;
            try
            {
                Guid guid = Guid.Empty;
                CLSIDFromProgID("Outlook.Application", out guid);
                if (guid != Guid.Empty)
                {
                    IDispatch pApp = null;
                    if (Process.GetProcessesByName("OUTLOOK").Count() > 0)
                    {
                        // Not needed as Outlook refreshes Inbox/Outbox automatically when it is opened

                        //IntPtr pUnk = IntPtr.Zero;
                        //HRESULT hr = GetActiveObject(ref guid, IntPtr.Zero, out pUnk);
                        //if (hr == HRESULT.S_OK)
                        //{
                        //    pApp = Marshal.GetObjectForIUnknown(pUnk) as IDispatch;                           
                        //}
                        bOutlookRunning = true;
                    }
                    else
                    {
                        pApp = (IDispatch)Activator.CreateInstance(Type.GetTypeFromCLSID(guid));
                    }
                                            
                    if (pApp != null)
                    {
                        pOutlook = (_Application)pApp;
                        //string sName;
                        //pOutlook.get_Name(out sName);                       
                        pOutlook.GetNamespace("MAPI", out ns);
                        if (ns != null)
                        {                          
                            ns.Logon("", "", false, Missing.Value);
                            ns.SendAndReceive((cbOutlookRefresh.IsChecked == true) ? true : false);
                        }                      
                    }
                }
            }
            catch (Exception ex)
            {
              
            }
            finally
            {
                if (!bOutlookRunning)
                {
                    if (ns != null)
                        Marshal.ReleaseComObject(ns);
                    if (pOutlook != null)
                    {
                        pOutlook.Quit();
                        Marshal.ReleaseComObject(pOutlook);
                    }
                }
            }
            return System.Threading.Tasks.Task.CompletedTask;
        }

        private TimeSpan tsDuration;
        private DateTime tsEnd;

        private void StartTimer(int nMilliSeconds, int nInterval)
        {
            dTimer = new Microsoft.UI.Xaml.DispatcherTimer();
            dTimer.Interval = TimeSpan.FromMilliseconds(nInterval);
            tsDuration = TimeSpan.FromMilliseconds(nMilliSeconds);
            tsEnd = DateTime.UtcNow + tsDuration;
            dTimer.Tick += Dt_Tick;
            dTimer.Start();
        }

        private void Dt_Tick(object sender, object e)
        {
            DateTime dtNow = DateTime.UtcNow;
            if (dtNow >= tsEnd)
            {               
                if (dTimer != null)
                {
                    dTimer.Stop();
                    dTimer = null;                   
                }
                bool isQueued = this.DispatcherQueue.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Normal, async () =>
                {
                   await RefreshOutlook();
                });               
            }
            else
            {
                
            }
        }

        private void btnShare_Click(object sender, RoutedEventArgs e)
        {
            HRESULT hr = DataTransferManagerInterop.ShowShareUIForWindow(hWnd);
        }

        // For bug on ContextMenu on (Rich)TextBox controls
        //https://github.com/microsoft/microsoft-ui-xaml/issues/4804
        private void Menu_Opening(object sender, object e)
        { 
            var tcbf = (Microsoft.UI.Xaml.Controls.TextCommandBarFlyout)sender;
            tcbf.Hide();
            //tcbf.ShowMode = FlyoutShowMode.TransientWithDismissOnPointerMoveAway;
            return;
        }
    }
}
