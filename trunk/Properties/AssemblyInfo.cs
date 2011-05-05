using System.Reflection;
using System.Runtime.InteropServices;

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("DragForIE9")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("s.weyl")]
[assembly: AssemblyProduct("DragForIE9")]
[assembly: AssemblyCopyright("Copyright © Microsoft 2010")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// Setting ComVisible to false makes the types in this assembly not visible 
// to COM components.  If you need to access a type in this assembly from 
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("01909860-75ef-47c6-88f0-329d66f1c7ed")]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Build and Revision Numbers 
// by using the '*' as shown below:
// [assembly: AssemblyVersion("1.0.*")]
[assembly: AssemblyVersion("0.9.2.1")]
[assembly: AssemblyFileVersion("0.9.2.1")]

/*
 * 0.1 初步构成
 * 2 对刷新进行了判断
 * 3 改用延时对刷新进行了判断
 * 4 改用IHTMLDocument2和IHTMLDocument3之间的转换，使事件的转换不丢失数据
 * 5 对新建网页是否已经设定BHO进行了改写，定为beta
 * 6 对刷新和导航的几种情况，采用计时判断
 * */
