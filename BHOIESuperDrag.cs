/****************************** Module Header ******************************\
* Module Name:  BHOIEContextMenu.cs
* Project:	    CSBrowserHelperObject
* Copyright (c) Microsoft Corporation.
* 
* The class BHOIEContextMenu is a Browser Helper Object which runs within Internet
* Explorer and offers additional services.
* 
* A BHO is a dynamic-link library (DLL) capable of attaching itself to any new 
* instance of Internet Explorer or Windows Explorer. Such a module can get in touch 
* with the browser through the container's site. In general, a site is an intermediate
* object placed in the middle of the container and each contained object. When the
* container is Internet Explorer (or Windows Explorer), the object is now required 
* to implement a simpler and lighter interface called IObjectWithSite. 
* It provides just two methods SetSite and GetSite. 
* 
* This class is used to disable the default context menu in IE. It also supplies 
* functions to register this BHO to IE.
* 
* This source is subject to the Microsoft Public License.
* See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
* All other rights reserved.
* 
* THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
* EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
* WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

/*
 * This code is based on "CSBrowserHelperObject" publish by Microsoft Corporation.
 * The origianl words are kept with out any modifications as above. 
 */


using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using SHDocVw;
using mshtml;
using System.Windows.Forms;
using System.Text;



namespace BHOForIE9
{
    /// <summary>
    /// Set the GUID of this class and specify that this class is ComVisible.
    /// A BHO must implement the interface IObjectWithSite. 
    /// </summary>
    [ComVisible(true),
    ClassInterface(ClassInterfaceType.None),
   Guid("3640b4a2-20aa-4f05-a575-b2ea1d7452e1")]

    public class SuperDrag :
        IObjectWithSite
    {
        // Current IE instance. For IE7 or later version, an IE Tab is just 
        // an IE instance.

        #region Global parameters
        public InternetExplorer ieInstance;
        public IHTMLDocument3 document;
        public HTMLElementEvents2_Event rootElementEvents = null;
        public int preY;
        public bool Refreshed = false;

        // To register a BHO, a new key should be created under this key.
        private const string BHORegistryKey =
            "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Browser Helper Objects";

        public const string ConfigKey = "Software\\s.weyl\\DragForIE9";

        public int NewTabGround;
        public string SearchString;        

        #endregion

        #region Com Register/UnRegister Methods
        /// <summary>
        /// When this class is registered to COM, add a new key to the BHORegistryKey 
        /// to make IE use this BHO.
        /// On 64bit machine, if the platform of this assembly and the installer is x86,
        /// 32 bit IE can use this BHO. If the platform of this assembly and the installer
        /// is x64, 64 bit IE can use this BHO.
        /// </summary>
        [ComRegisterFunction]
        public static void RegisterBHO(Type t)
        {
            RegistryKey key = Registry.LocalMachine.OpenSubKey(BHORegistryKey, true);
            if (key == null)
            {
                key = Registry.LocalMachine.CreateSubKey(BHORegistryKey);
            }

            // 32 digits separated by hyphens, enclosed in braces: 
            // {00000000-0000-0000-0000-000000000000}
            string bhoKeyStr = t.GUID.ToString("B");

            RegistryKey bhoKey = key.OpenSubKey(bhoKeyStr, true);

            // Create a new key.
            if (bhoKey == null)
            {
                bhoKey = key.CreateSubKey(bhoKeyStr);
            }

            // NoExplorer:dword = 1 prevents the BHO to be loaded by Explorer
            string name = "NoExplorer";
            object value = (object)1;
            bhoKey.SetValue(name, value);
            key.Close();
            bhoKey.Close();

            key = Registry.CurrentUser.OpenSubKey(ConfigKey, true);
            if (key == null)
            {
                key = Registry.CurrentUser.CreateSubKey(ConfigKey);
                name = "NewTabGround";
                value = (object)1;
                key.SetValue(name, value);
                name = "SearchString";
                value = (object)"http://www.google.com.hk/search?hl=zh-CN&q={0}";
                key.SetValue(name, value);
            }
            key.Close();
        }

        /// <summary>
        /// When this class is unregistered from COM, delete the key.
        /// </summary>
        [ComUnregisterFunction]
        public static void UnregisterBHO(Type t)
        {
            RegistryKey key = Registry.LocalMachine.OpenSubKey(BHORegistryKey, true);
            string guidString = t.GUID.ToString("B");
            if (key != null)
            {
                key.DeleteSubKey(guidString, false);
            }

            string BHOMenuRegistryKey = "Software\\Microsoft\\Internet Explorer\\Extensions";
            RegistryKey key1 = Registry.CurrentUser.OpenSubKey(BHOMenuRegistryKey, true);
            guidString = "{aa4a9427-8721-4651-8d6b-a25e623c8bbd}";
            if (key1 != null)
            {
                key1.DeleteSubKey(guidString, false);
            }

            Registry.CurrentUser.DeleteSubKeyTree("Software\\s.weyl", false);


        }
        #endregion

        #region IObjectWithSite Members
        /// <summary>
        /// This method is called when the BHO is instantiated and when
        /// it is destroyed. The site is an object implemented the 
        /// interface InternetExplorer.
        /// </summary>
        /// <param name="site"></param>
        public void SetSite(Object site)
        {
            if (site != null)
            {
                ieInstance = (InternetExplorer)site;

                /*ieInstance.BeforeNavigate2 +=
                    new DWebBrowserEvents2_BeforeNavigate2EventHandler(
                        ieInstance_BeforeNavigate2);*/
                ieInstance.NavigateComplete2 +=
                    new DWebBrowserEvents2_NavigateComplete2EventHandler(
                        ieInstance_NavigateComplete2);
                ieInstance.DownloadBegin +=
                    new DWebBrowserEvents2_DownloadBeginEventHandler(
                        ieInstance_DownloadBegin);
                ieInstance.DownloadComplete +=
                    new DWebBrowserEvents2_DownloadCompleteEventHandler(
                        ieInstance_DownloadComplete);

                RegistryKey key = Registry.CurrentUser.OpenSubKey(ConfigKey, false);
                if (key != null)
                {
                    NewTabGround = Convert.ToInt32(key.GetValue("NewTabGround"));
                    SearchString = key.GetValue("SearchString").ToString();
                }
                else
                {
                    NewTabGround = 1;
                    SearchString = "http://www.google.com.hk/search?hl=zh-CN&q={0}";
                }
                key.Close();
            }

        }

        /// <summary>
        /// Retrieves and returns the specified interface from the last site
        /// set through SetSite(). The typical implementation will query the
        /// previously stored pUnkSite pointer for the specified interface.
        /// </summary>
        public void GetSite(ref Guid guid, out Object ppvSite)
        {
            IntPtr punk = Marshal.GetIUnknownForObject(ieInstance);
            ppvSite = new object();
            IntPtr ppvSiteIntPtr = Marshal.GetIUnknownForObject(ppvSite);
            int hr = Marshal.QueryInterface(punk, ref guid, out ppvSiteIntPtr);
            Marshal.ThrowExceptionForHR(hr);
            Marshal.Release(punk);
            Marshal.Release(ppvSiteIntPtr);
        }
        #endregion

        #region WebbrowserEvent

        /// <summary>
        /// Handle the DocumentComplete event.
        /// </summary>
        /// <param name="pDisp">
        /// The pDisp is an an object implemented the interface InternetExplorer.
        /// By default, this object is the same as the ieInstance, but if the page 
        /// contains many frames, each frame has its own document.
        /// </param>
        void ieInstance_NavigateComplete2(object pDisp, ref object URL)
        {
            if (pDisp == ieInstance)
            {
                if (Refreshed == false)
                {
                    if (URL.ToString().StartsWith("http"))
                    {
                        ieInstance.Navigate2(URL.ToString());
                    }
                    Refreshed = true;
                }
                else
                {
                    document = ieInstance.Document as IHTMLDocument3;
                    rootElementEvents = document.documentElement as HTMLElementEvents2_Event;
                    SetDragHandler();
                }
            }
        }


        void ieInstance_DocumentComplete(object pDisp, ref object URL)
        {

        }


        void ieInstance_DownloadBegin()
        {
            if (rootElementEvents != null)
            {
                SetDragHandler();
            }
            else
            {
                document = ieInstance.Document as IHTMLDocument3;
                rootElementEvents = document.documentElement as HTMLElementEvents2_Event;
                SetDragHandler();
            }
        }


        void ieInstance_DownloadComplete()
        {

        }

        #endregion

        #region HTMLElementEvents event
        /// <summary>
        /// Handle the HTMLElementEvents event.
        /// </summary>
        /// <param name="e"></param>
        /// 

        void SetDragHandler()
        {

            //To avoid set handler repeatedly, remove previous handler.
            rootElementEvents.ondragstart -=
                new HTMLElementEvents2_ondragstartEventHandler(
                    Events_Ondragstart);
            rootElementEvents.ondragstart +=
                new HTMLElementEvents2_ondragstartEventHandler(
                    Events_Ondragstart);
            rootElementEvents.ondragover += (e) => false;
            rootElementEvents.ondragend -=
                new HTMLElementEvents2_ondragendEventHandler(
                    Events_Ondragend);
            rootElementEvents.ondragend +=
                new HTMLElementEvents2_ondragendEventHandler(
                     Events_Ondragend);
        }

        bool Events_Ondragstart(IHTMLEventObj e)
        {
            preY = e.clientY;
            return true;
        }

        void Events_Ondragend(IHTMLEventObj e)
        {
            BrowserNavConstants n = BrowserNavConstants.navOpenInBackgroundTab;
            switch (NewTabGround)
            {
                case 1:
                    n = BrowserNavConstants.navOpenInBackgroundTab;
                    break;
                case 2:
                    n = BrowserNavConstants.navOpenNewForegroundTab;
                    break;
                case 3:
                    if (e.clientY < preY)
                        n = BrowserNavConstants.navOpenNewForegroundTab;
                    else
                        n = BrowserNavConstants.navOpenInBackgroundTab;
                    break;
                case 4:
                    if (e.clientY >= preY)
                        n = BrowserNavConstants.navOpenNewForegroundTab;
                    else
                        n = BrowserNavConstants.navOpenInBackgroundTab;
                    break;
            }

            var eventObj = e as IHTMLEventObj2;

            //When drag a url.
            var url = (object)eventObj.dataTransfer.getData("URL") as string;
            if (!string.IsNullOrEmpty(url))
            {
                ieInstance.Navigate2(url, n);
                return;
            }

            //When drag a text.
            var text = (object)eventObj.dataTransfer.getData("TEXT") as string;
            if (!string.IsNullOrEmpty(text))
            {
                if (text.StartsWith("http://") || text.StartsWith("https://"))
                {
                    ieInstance.Navigate2(text, n);
                }
                else
                {
                    ieInstance.Navigate2(string.Format(SearchString, text), n);
                }
                return;
            }
        }

        #endregion

    }
}
