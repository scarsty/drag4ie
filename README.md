# cs-drag4ie9
为IE9以后的版本增加超级拖拽功能

Super Drag and Go (Beta) for IE
超级拖拽(beta) for IE

现在可以在Windows 8上的IE10，以及Windows7 上的IE10RP版中使用。 我没有测试过更多平台，如果有问题请联系我。

注意需要 .NET framework 4 的完全版，一般来说操作系统自带的是一个基本版。请在安装指引里面找下载的链接。

## 简介

### **1. 如何安装和卸载**

首先需要你的系统中安装了.NET Framework 4.0，可以在 <http://www.microsoft.com/downloads/en/details.aspx?familyid=9cfb2d51-5ff4-4491-b0e5-b386f32c0992>下载。

*手动安装（推荐）：*

下载压缩包，解压到任意目录，以管理员身份运行“install.cmd”。如果要在64位IE中使用，运行“install64.cmd"。

卸载时运行“uninstall.cmd”，这会将32位和64位版本同时移除。

“CustomDrag.exe”这个程序可以对拖拽进行一些设置，还可以将自身添加到IE的工具菜单中。这些设置在卸载时都会被删除。

### **2. 设置搜索字串**

用你喜欢的搜索引擎搜索“{0}”，然后把地址栏中的内容复制下来即可。 需要注意的是，如果这个字串中没有出现“{0}”，而是“%7B0%7D”，则需要将“%7B0%7D”替换为“{0}”。

### **3. 参考文献**

<http://www.cnblogs.com/TianFang/archive/2011/03/12/1982100.html>

<http://www.codeproject.com/KB/cs/Issuewithbandobjects.aspx>

MSDN.

以及其他。

## Introduction in English

### **1. How to install and uninstall**

First of all, please make sure you have .NET Framework 4.0. You can get it from <http://www.microsoft.com/downloads/en/details.aspx?familyid=9cfb2d51-5ff4-4491-b0e5-b386f32c0992>.

*Manually:*

To install, you have to decompress the ZIP file into any directory, then run "install.cmd" or "install64.cmd" (for 64-bit IE) as Administrator.

To uninstall, run "uninstall.cmd" which will remove both 32-bit and 64-bit versions.

You can also use "CustomDrag.exe" to modify the behavior of this BHO, or add a link of this program into IE tools menu. All settings will be deleted when uninstall.

Hope you can like it.

### **2. The "Search string"**

Please search "{0}" in the engineering you prefer, then just copy the string in the address bar.

But you have to note, if there are not "{0}" in the string, but "%7B0%7D" instead, please change "%7B0%7D" to "{0}", then use it as "Search string".

### **3. References**

<http://www.cnblogs.com/TianFang/archive/2011/03/12/1982100.html>

<http://www.codeproject.com/KB/cs/Issuewithbandobjects.aspx>

Contents on MSDN.

And so on, many thanks to them.

Seek setup files in: <http://code.google.com/p/cs-drag4ie9/downloads/list>



## 更新记录

- 0.1 初步构成
- 0.2 对刷新进行了判断
- 0.3 改用延时对刷新进行了判断
- 0.4 改用IHTMLDocument2和IHTMLDocument3之间的转换，使事件的转换不丢失数据
- 0.5 对新建网页是否已经设定BHO进行了改写，定为beta
- 0.6 对刷新和导航的几种情况，采用计时判断
- 0.9 网页元素事件在主文件直接处理；在每次DownloadBegin时重新注入事件，增加事件响应先移除原有的响应；增加一个设置用程序
- 0.9.1 在新选项卡重新导航一次，使得事件可以被注入
- 0.9.2 增加一个安装包，添加Installer类
- 0.9.2.1 当新选项卡打开的不是http页面时不重新导航
- 0.9.2.2 刷新之后仍然有效
- 0.9.2.3 将搜索字串转为URL用的格式
- 0.9.5 原本对新选项卡的理解有误，现改为只注入一次，不重新导航；意外发现判断刷新的方法，因此每个网页只注入一到二次
- 0.9.6 更新了获取字串的时机，可以在IE10中使用



## C#编写IE插件的一些经验

### **1. BeforeNavigate2不发生**

在C#中这个事件好像是有问题，一直没有被触发过。这个有可能是.NET Framework 4.0的bug，因为在C++中调用这个事件是有效的。

除了等微软发补丁，我们自己没什么解决办法。

这个问题似乎是解决了（后注）。

### **2. 多个NavigateComplete2和DocumentComplete的问题**

这个问题的一般结论是，在载入多个框架的网页时，每个框架都会引发自己的NavigateComplete2和DocumentComplete事件，判断该事件是否为主框架的对应事件可以用以下代码： 
```c#
void ieInstance_NavigateComplete2(object pDisp, ref object URL) 
{ if (pDisp == ieInstance) { //Do something. } } 
```
其中ieInstance为public对象，而SetSite中将其指定： 
```c#
public InternetExplorer ieInstance;
ieInstance = (InternetExplorer)site; 
```
可以认为是该IE窗口（或选项卡）对应的对象。

但是在某些复杂页面上，该事件似乎会被触发多次。在大部分网页上应该是只触发一次的。

### **3. 如何判断点击刷新按钮事件**

其实我觉得Chrome的作法比较奇怪，在网页加载过程是不能刷新的，这个细节非常有趣。

因为我编写的BHO主要是设置网页元素事件，而这类事件会在刷新之后失效，因此我需要捕捉按刷新页面这个事件。

网页刷新过程中，NavigateComplete2和DocumentComplete都是不会发生的。同时IE并未提供这个按钮的事件，因此有很多不太正规的方法去判断。MSDN上说的是，如果一个DownloadBegin之前是DocumentComplete，那么该DownloadBegin应该是由刷新按钮触发的。

这个方法是不正确的，事实上我已经观测到很多网页在加载的过程中，DocumentComplete之后还有DownloadBegin。这多半是因为页面上有一些额外的元素，例如漂浮的图片广告之类。同时，该方法对于在加载过程中刷新，或者是刷新之后再刷新都是无效的。

我曾经用过一个方法，基于一个可能性较大的猜测：DownloadBegin必然有接下来的DownloadComplete，即使操作被挂起。那么，DownloadComplete之后呢？一般来说应该是紧接着一个新的DownloadBegin，如果没有，就是网页加载完毕了。据此，可以记录下每个DownloadComplete发生的时间，如果某一个DownloadBegin发生时，在其之前一定时间之内没有任何事件发生，即可以认为这个DownloadBegin是由刷新按钮触发的。

但是这个方法也有相当的缺陷，首先是间隔时间长度，设得太短容易误判，太长容易没反应，毕竟是基于人们浏览网页的习惯。此外，有一种特殊情况，就是在网页加载过程中刷新，很明显这种方法是不行的。

不过后来很意外，我发现网页元素在刷新后失效这个特性可以利用一下，具体的方法如下：

定义public变量： 
```c#
public IHTMLDocument3 document; 
public HTMLElementEvents2_Event rootElementEvents = null;
```
在NavigateComplete2中将网页元素事件指定： 
```c#
document = ieInstance.Document as IHTMLDocument3; 
rootElementEvents = document.documentElement as HTMLElementEvents2_Event; 
```
在DownloadBegin中添加如下代码：
```c#
if (rootElementEvents != document.documentElement as HTMLElementEvents2_Event) 
{ //This might be refreshing, if no navigations. } 
```
该if语句就是判断rootElementEvents是否跟ieInstance挂钩，刷新之后既然不挂钩了，那么该判断为真时，就可以认为是用户刷新了网页。而为了使该判断仍然可以使用，之后需要在立刻在下一个DownloadComplete重新将二者挂钩。

该方法对任何时候的刷新都有效，但是有时在打开新页面时也会发生，可以通过一些小手段处理。

### **4. 避免网页元素事件的重复设定**

用户打开一个选项卡之后，所进行的操作是难以估计的，有可能是重新打开该页面，刷新该页面（以上二者是不同的），从该选项卡打开新页面等等。一般来说，我们需要每次重新导航或者刷新之后，重新设定网页元素的事件。

正是因为IE没有刷新事件的接口，因此重新设定的适合时机不太容易抓住，前面所说的判断刷新的方法虽然有效，但是却不太容易将它与新的导航区分开来。而刷新和重导航是都需要设定事件的，这时就可能发生重复设定。

如果需要设定的事件并没有什么明显的操作，那很多时候倒也无关痛痒。但是如果操作是可见的，那么重复设定的后果就是可见操作会发生两次甚至更多次。例如在本例中，需要对鼠标的拖拽事件进行设定，可以拖拽出新的选项卡，那么重复设定就会导致一次拖拽出现两个甚至更多个新选项卡，这显然是我们不愿意看到的。

如果实在太难以区分，那么就干脆不要区分而采用其他的手段。在设定事件之前，首先清除前一个事件，即：

```c#
rootElementEvents.ondragend -= new HTMLElementEvents2_ondragendEventHandler( Events_Ondragend);
rootElementEvents.ondragend += new HTMLElementEvents2_ondragendEventHandler( Events_Ondragend); 
```
在设定之前首先清除，即可以保证事件只被设定一次，rootElementEvents应该是一个public的全局变量，而不是单属于这个函数。因为-=操作在相关事件没有设定的时候是不做操作，所以首次设定也不会有错误发生，该方法被证实是比较有效的。
