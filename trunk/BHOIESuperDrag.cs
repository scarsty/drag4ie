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


using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using SHDocVw;
using mshtml;
using System.Windows.Forms;


namespace CSBHODragForIE9
{
    /// <summary>
    /// Set the GUID of this class and specify that this class is ComVisible.
    /// A BHO must implement the interface IObjectWithSite. 
    /// </summary>
    [ComVisible(true),
    ClassInterface(ClassInterfaceType.None),
   Guid("c963e2da-e71f-4dbe-9fcd-f252075b0aa7")]
    public class BHOIESuperDrag : IObjectWithSite
    {
        // Current IE instance. For IE7 or later version, an IE Tab is just 
        // an IE instance.
        public InternetExplorer ieInstance;

        public bool SetHandlered = true, DocumentLoaded = false, Navigated = false, Refreshed = false, NewPage = false;
        public DateTime timer = DateTime.Now.AddDays(-1), timerPreSet = DateTime.Now.AddDays(-1),
            timerPreNav = DateTime.Now.AddDays(-1), timerPreRef = DateTime.Now.AddDays(-1);
        //public HTMLDocumentEventHelper helper;
        public string urlPreNavigate = "";
        public IHTMLDocument3 document;
        public HTMLElementEvents2_Event rootElementEvents;

        // To register a BHO, a new key should be created under this key.
        private const string BHORegistryKey =
            "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Browser Helper Objects";



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
            //HTMLDocument doc1 = this.ieInstance.Document as HTMLDocument;
            if (site != null)
            {
                ieInstance = (InternetExplorer)site;


                ieInstance.BeforeNavigate2 +=
                    new DWebBrowserEvents2_BeforeNavigate2EventHandler(
                        ieInstance_BeforeNavigate2);
                ieInstance.NavigateComplete2 +=
                    new DWebBrowserEvents2_NavigateComplete2EventHandler(
                        ieInstance_NavigateComplete2);
                ieInstance.DocumentComplete -=
                    new DWebBrowserEvents2_DocumentCompleteEventHandler(
                        ieInstance_DocumentComplete);
                ieInstance.DocumentComplete +=
                    new DWebBrowserEvents2_DocumentCompleteEventHandler(
                        ieInstance_DocumentComplete);
                ieInstance.DownloadComplete +=
                    new DWebBrowserEvents2_DownloadCompleteEventHandler(
                        ieInstance_DownloadComplete);
                ieInstance.DownloadBegin +=
                    new DWebBrowserEvents2_DownloadBeginEventHandler(
                        ieInstance_DownloadBegin);


                //document = ieInstance.Document as IHTMLDocument3;

                //var rootElementEvents = document.documentElement as HTMLElementEvents_Event;

                /*rootElementEvents.ondragstart += 
                    new HTMLElementEvents2_ondragstartEventHandler(
                        Events_Ondragstart);
                //rootElementEvents.ondragover += () => false;
                rootElementEvents.ondrop +=
                    new HTMLElementEvents_ondropEventHandler(
                        Events_Ondrop);*/
                document = ieInstance.Document as IHTMLDocument3;
                rootElementEvents = document.documentElement as HTMLElementEvents2_Event;
                rootElementEvents.oncopy +=
                    new HTMLElementEvents2_oncopyEventHandler(
                        Events_Oncopy);
                //ieInstance.ProgressChange +=
                //new DWebBrowserEvents2_ProgressChangeEventHandler(
                //ieInstance_ProgressChange);


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

        #region event handler

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

            Navigated = true;
            //MessageBox.Show("navigate"+URL.ToString());

            if (pDisp == ieInstance)
            {
                //MessageBox.Show(Events_Ondrop(null).ToString());
                document = ieInstance.Document as IHTMLDocument3;
                rootElementEvents = document.documentElement as HTMLElementEvents2_Event;
                rootElementEvents.ondragstart +=
                    new HTMLElementEvents2_ondragstartEventHandler(
                        Events_Ondragstart);
                rootElementEvents.ondragover += (e) => false;
                rootElementEvents.ondragend +=
                    new HTMLElementEvents2_ondragendEventHandler(
                        Events_Ondragend);
                rootElementEvents.ondrop +=
                    new HTMLElementEvents2_ondropEventHandler(
                        Events_Ondrop);
                //MessageBox.Show("main navigate"+URL.ToString());
                if ((urlPreNavigate != URL.ToString()) ||
                    (PassedMilliseconds(timerPreNav) > 5000) && ((urlPreNavigate == URL.ToString())))
                {
                    //if (Refreshed == false)

                    //MessageBox.Show("navigate");
                    if (PassedMilliseconds(timerPreRef) > 1000)
                    {
                        NewPage = true;
                        TrySetHandler();
                        //MessageBox.Show("navigate and set 1");
                        timerPreNav = DateTime.Now;
                    }
                    else
                    {
                        if (NewPage == true)
                        {
                            TrySetHandler();
                            //MessageBox.Show("navigate and set 2");
                        }
                    }

                }


                urlPreNavigate = URL as string;
                DocumentLoaded = false;
            }

            //timer = DateTime.Now;

        }


        void ieInstance_DocumentComplete(object pDisp, ref object URL)
        {

            if (pDisp == ieInstance)
            {
                DocumentLoaded = true;

                //MessageBox.Show("docuemnt complete" + URL.ToString());


                //if (Navigated == false)
                {
                    //NewPage = true;
                    //TrySetHandler();
                }

                //if (Navigated == true)
                //Navigated = false;
                Refreshed = false;

                timerPreNav = DateTime.Now;
                urlPreNavigate = URL as string;

            }

            timer = DateTime.Now;

        }

        void ieInstance_DocumentComplete2(object pDisp, ref object URL)
        {
            MessageBox.Show("^");
        }


        void ieInstance_DownloadBegin()
        {

            rootElementEvents = document.documentElement as HTMLElementEvents2_Event;
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
            
            

            if (PassedMilliseconds(timer) > 5000)
            {
                //MessageBox.Show("refresh");

                if (PassedMilliseconds(timerPreNav) > 1000)
                {
                    NewPage = true;
                    object purl = urlPreNavigate;
                    //ieInstance_NavigateComplete2(ieInstance, ref purl);
                    //MessageBox.Show("refresh real");
                }
            }

        }


        void ieInstance_DownloadComplete()
        {

            /*if (NewPage)
            {
                TrySetHandler();
                //MessageBox.Show("Download Complete and set");

            }*/
            timer = DateTime.Now;

        }


        void ieInstance_BeforeNavigate2(object pDisp, ref object URL, ref object Flags,
            ref object TargetFrameName, ref object PostData, ref object Headers, ref bool Cancel)
        {
            Navigated = true;
        }


        void TrySetHandler()
        {
            if (NewPage && PassedMilliseconds(timerPreSet) > 2000)
            {
                if (ieInstance != null)
                {
                    SetHandler(ieInstance);
                }
                timerPreSet = DateTime.Now;
                NewPage = false;
            }


        }

        double PassedMilliseconds(DateTime d)
        {
            TimeSpan s = DateTime.Now - d;
            return s.TotalMilliseconds;
        }

        /*void ieInstance_ProgressChange(int Progress, int ProgressMax)
        {
            if (DownloadBegined == false && isRefresh == false && Progress < ProgressMax && Progress > 0)
            {
                isRefresh = true;
                MessageBox.Show(Progress.ToString() + "might in refresh" + DownloadBegined.ToString() + isRefresh.ToString());

            }
            if (DownloadBegined == false && isRefresh == true && Progress == -1)
            {
                isRefresh = false;
            }
        }*/


        void SetHandler(InternetExplorer explorer)
        {
            try
            {
                // Register the oncontextmenu event of the  document in InternetExplorer.
                //MessageBox.Show("set");
                //HTMLDocumentEventHelper helper =
                //new HTMLDocumentEventHelper(ieInstance.Document as IHTMLDocument3, explorer);
                //helper.ondragstart += new HtmlEvent(ondragstartHandler);
                //var document = ieInstance.Document as IHTMLDocument3;
                //var rootElementEvents = document.documentElement as HTMLElementEvents_Event;
                //rootElementEvents.ondragstart += new HtmlEvent(ondragstartHandler); //() => false;
                //rootElementEvents.ondragover += () => false;
                //rootElementEvents.ondrop += () => { SuperDragDrop(); return false; };

                //helper.ondragstart -= new HtmlEvent(ondragstartHandler);
                //SetHandlered = true;
                //CommandSet = false;
            }
            catch { }
        }

        /// <summary>
        /// Handle the oncontextmenu event.
        /// </summary>
        /// <param name="e"></param>
        bool Events_Ondragstart(IHTMLEventObj e)
        {

            // To cancel the default behavior, set the returnValue property of the event
            // object to false.
            //this.ondragstartHandler += e => e.returnValue = false;
            //MessageBox.Show("dd");
            return true;

        }

        bool Events_Ondrop(IHTMLEventObj e)
        {

            return false;
        }

        void Events_Ondragend(IHTMLEventObj e)
        {
            superdrog();
        }

        bool Events_Oncopy(IHTMLEventObj e)
        {
            //MessageBox.Show(e.ToString());
            return false;
        }

        #endregion

        #region

        void superdrog()
        {
            var doc1 = ieInstance.Document as IHTMLDocument2;
            var eventObj = doc1.parentWindow.@event as IHTMLEventObj2;

            //拖拽的是链接，在新窗口中打开链接
            var url = (object)eventObj.dataTransfer.getData("URL") as string;
            //MessageBox.Show(url);
            if (!string.IsNullOrEmpty(url))
            {
                //MessageBox.Show(url);                 
                ieInstance.Navigate2(url, BrowserNavConstants.navOpenInBackgroundTab);

                return;
            }

            //拖拽的是选择的文本，则用google搜索该文本
            var text = (object)eventObj.dataTransfer.getData("TEXT") as string;
            if (!string.IsNullOrEmpty(text))
            {
                if (text.StartsWith("http://", System.StringComparison.OrdinalIgnoreCase))    //未被识别的超链接
                {
                    ieInstance.Navigate2(text, BrowserNavConstants.navOpenInBackgroundTab);
                }
                else    //待搜索的文本
                {
                    ieInstance.Navigate2(string.Format("http://www.google.com.hk/search?hl=zh-CN&q={0}", text), BrowserNavConstants.navOpenInBackgroundTab);
                }
                return;
            }
            return;
        }

        #endregion

    }
}
