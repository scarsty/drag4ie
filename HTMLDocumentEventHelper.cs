/****************************** Module Header ******************************\
* Module Name:  HTMLDocumentEventHelper.cs
* Project:	    CSBrowserHelperObject
* Copyright (c) Microsoft Corporation.
* 
* This ComVisible class HTMLDocumentEventHelper is used to set the event handler
* of the HTMLDocument. The interface DispHTMLDocument defines many events like 
* oncontextmenu, onclick and so on, and these events could be set to an
* HTMLEventHandler instance.
* 
* 
* This source is subject to the Microsoft Public License.
* See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
* All other rights reserved.
* 
* THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
* EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
* WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

using System.Runtime.InteropServices;
using SHDocVw;
using mshtml;
using System;
using System.Windows.Forms;
using System.Threading;

namespace CSBHODragForIE9
{
    [ComVisible(true)]

    [Flags]
    public enum BrowserNavConstants
    {
        /// <summary>
        /// Open the resource or file in a new window.
        /// </summary>
        navOpenInNewWindow = 0x1,
        /// <summary>
        /// Do not add the resource or file to the history list. The new page replaces the current page in the list.
        /// </summary>
        navNoHistory = 0x2,
        /// <summary>
        /// Do not consult the Internet cache; retrieve the resource from the origin server (implies BINDF_PRAGMA_NO_CACHE and BINDF_RESYNCHRONIZE).
        /// </summary>
        navNoReadFromCache = 0x4,
        /// <summary>
        /// Do not add the downloaded resource to the Internet cache. See BINDF_NOWRITECACHE.
        /// </summary>
        navNoWriteToCache = 0x8,
        /// <summary>
        /// If the navigation fails, the autosearch functionality attempts to navigate common root domains (.com, .edu, and so on). If this also fails, the URL is passed to a search engine.
        /// </summary>
        navAllowAutosearch = 0x10,
        /// <summary>
        /// Causes the current Explorer Bar to navigate to the given item, if possible.
        /// </summary>
        navBrowserBar = 0x20,
        /// <summary>
        /// Microsoft Internet Explorer 6 for Microsoft Windows XP Service Pack 2 (SP2) and later. If the navigation fails when a hyperlink is being followed, this constant specifies that the resource should then be bound to the moniker using the BINDF_HYPERLINK flag.
        /// </summary>
        navHyperlink = 0x40,
        /// <summary>
        /// Internet Explorer 6 for Windows XP SP2 and later. Force the URL into the restricted zone.
        /// </summary>
        navEnforceRestricted = 0x80,
        /// <summary>
        /// Internet Explorer 6 for Windows XP SP2 and later. Use the default Popup Manager to block pop-up windows.
        /// </summary>
        navNewWindowsManaged = 0x0100,
        /// <summary>
        /// Internet Explorer 6 for Windows XP SP2 and later. Block files that normally trigger a file download dialog box.
        /// </summary>
        navUntrustedForDownload = 0x0200,
        /// <summary>
        /// Internet Explorer 6 for Windows XP SP2 and later. Prompt for the installation of Microsoft ActiveX controls.
        /// </summary>
        navTrustedForActiveX = 0x0400,
        /// <summary>
        /// Windows Internet Explorer 7. Open the resource or file in a new tab. Allow the destination window to come to the foreground, if necessary.
        /// </summary>
        navOpenInNewTab = 0x0800,
        /// <summary>
        /// Internet Explorer 7. Open the resource or file in a new background tab; the currently active window and/or tab remains open on top.
        /// </summary>
        navOpenInBackgroundTab = 0x1000,
        /// <summary>
        /// Internet Explorer 7. Maintain state for dynamic navigation based on the filter string entered in the search band text box (wordwheel). Restore the wordwheel text when the navigation completes.
        /// </summary>
        navKeepWordWheelText = 0x2000
    } 

    public class HTMLDocumentEventHelper
    {
        private IHTMLDocument2 document;
        private InternetExplorer ieInstance;

        public HTMLDocumentEventHelper(IHTMLDocument3 document, InternetExplorer ieInstance)
        {
            this.document = document as IHTMLDocument2;
            this.ieInstance = ieInstance;

            this.ondragstart += e => e.returnValue = false;
            var rootElementEvents = document.documentElement as HTMLElementEvents_Event;
            rootElementEvents.ondragover += () => false;
            rootElementEvents.ondrop += () => { SuperDragDrop(); return false; };

        }

        public void SuperDragDrop()
        {
            
            //var doc1 = document as HTMLDocument;
            //Thread.Sleep(100);
            //MessageBox.Show("ddd");
            //var eventObj = doc1.parentWindow.@event as IHTMLEventObj2;
            var eventObj = document.parentWindow.@event as IHTMLEventObj2;
            
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

        public event HtmlEvent ondragstart
        {            
            
            add
            {
                

                object existingHandler = document.ondragstart;
                HTMLEventHandler handler = existingHandler as HTMLEventHandler;

                // Set the handler to the oncontextmenu event.
                //dispDoc.ondragstart = handler;

                if (handler != null) 
                    handler.eventHandler += value;
            }
            remove
            {
                //MessageBox.Show("remove");
                DispHTMLDocument dispDoc = this.document as DispHTMLDocument;
                object existingHandler = dispDoc.ondragstart;

                HTMLEventHandler handler = existingHandler is HTMLEventHandler ?
                    existingHandler as HTMLEventHandler : null;

                if (handler != null)
                    handler.eventHandler -= value;
            }
        }

        
    }

}
