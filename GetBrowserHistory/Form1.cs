using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;  //命名空间
using System.Reflection;               //提供加载类型 Pointer指针
using Microsoft.Win32;
using System.IO;                       //文件操作

namespace GetBrowserHistory
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();    //窗口初始化
        }
        public DateTime FileTimeToDatetime(System.Runtime.InteropServices.ComTypes.FILETIME ft)
        { 
            //定义函数，该函数接受一个FileTime类型的参数，返回一个DateTime类型的返回值。实现Filetime类型到DateTime类型的转换，便于输出
            long time;
            DateTime datetime;
            time = ((long)ft.dwHighDateTime << 32) + (long)ft.dwLowDateTime;  //将高位部分左移32位与低位部分组合
            datetime = DateTime.FromFileTime(time);           //调用DateTime类中的FromFileTime方法
            return datetime;
        }

        private void GetHistory()
        {
            IUrlHistoryStg2 _UrlHistoryStg2 = (IUrlHistoryStg2)new UrlHistory();    
            IEnumSTATURL _EnumSTATURL = _UrlHistoryStg2.EnumUrls();                 
            STATURL _STATURL;
            uint _Fectched;
            DateTime Datetime;
            int count = 0;

            FileStream NoteFile = new FileStream("C:\\Users\\oceanier\\Desktop\\BrowserHistory.txt", FileMode.Create, FileAccess.Write);//创建写入文件
            StreamWriter WriteNote = new StreamWriter(NoteFile);
            WriteNote.WriteLine("序号\t访问时间\t\t网址\n");
            while (_EnumSTATURL.Next(1, out _STATURL, out _Fectched) == 0)
            {
                if (_STATURL.pwcsUrl.Substring(0, 4) == "http"&&_STATURL.pwcsUrl.Length<=100)
                {
                    count += 1;
                    Datetime = FileTimeToDatetime(_STATURL.ftLastVisited);
                    richTextBox1.AppendText(string.Format("{0}\r\n{1}\r\n{2}\r\n", count, Datetime, _STATURL.pwcsUrl));
                    WriteNote.WriteLine(count+"\t"+Datetime+"\t"+_STATURL.pwcsUrl+"\n");
                }

            }
            MessageBox.Show("成功获取浏览器浏览记录！\n可在桌面txt文档中查看。", "提示");
        }
        private void button1_Click(object sender, EventArgs e)
        {
            GetHistory();
        }
    }
    #region COM接口实现获取IE历史记录
    //自定义结构 IUrlHistory
    public struct STATURL
    {
        public static uint SIZEOF_STATURL = (uint)Marshal.SizeOf(typeof(STATURL));
        public uint cbSize;                    //网页大小
        [MarshalAs(UnmanagedType.LPWStr)]
        public string pwcsUrl;     //网页Url
        [MarshalAs(UnmanagedType.LPWStr)]
        public string pwcsTitle;   //网页标题
        public System.Runtime.InteropServices.ComTypes.FILETIME ftLastVisited, ftLastUpdated, ftExpires;    //网页最近访问时间，网页最近更新时间
        public uint dwFlags;
    }

    //ComImport属性通过guid调用com组件
    [ComImport, Guid("3C374A42-BAE4-11CF-BF7D-00AA006946EE"),
        InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]

    interface IEnumSTATURL
    {
        [PreserveSig]
        //搜索IE历史记录匹配的搜索模式并复制到指定缓冲区
        uint Next(uint celt, out STATURL rgelt, out uint pceltFetched);
        void Skip(uint celt);
        void Reset();
        void Clone(out IEnumSTATURL ppenum);
        void SetFilter([MarshalAs(UnmanagedType.LPWStr)] string poszFilter, uint dwFlags);
    }

    [ComImport, Guid("AFA0DC11-C313-11d0-831A-00C04FD5AE38"),
        InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IUrlHistoryStg2
    {
        #region IUrlHistoryStg methods
        void AddUrl([MarshalAs(UnmanagedType.LPWStr)] string pocsUrl,
            [MarshalAs(UnmanagedType.LPWStr)] string pocsTitle, uint dwFlags);

        void DeleteUrl([MarshalAs(UnmanagedType.LPWStr)] string pocsUrl,
            uint dwFlags);

        void QueryUrl([MarshalAs(UnmanagedType.LPWStr)] string pocsUrl, uint dwFlags,
            ref STATURL lpSTATURL);

        void BindToObject([MarshalAs(UnmanagedType.LPWStr)] string pocsUrl, ref Guid riid,
            [MarshalAs(UnmanagedType.IUnknown)] out object ppvOut);

        IEnumSTATURL EnumUrls();
        #endregion

        void AddUrlAndNotify(
            [MarshalAs(UnmanagedType.LPWStr)] string pocsUrl,
            [MarshalAs(UnmanagedType.LPWStr)] string pocsTitle,
            uint dwFlags,
            [MarshalAs(UnmanagedType.Bool)] bool fWriteHistory,
            [MarshalAs(UnmanagedType.IUnknown)] object    /*IOleCommandTarget*/
            poctNotify,
            [MarshalAs(UnmanagedType.IUnknown)] object punkISFolder);

        void ClearHistory();       //清除历史记录
    }

    [ComImport, Guid("3C374A40-BAE4-11CF-BF7D-00AA006946EE")]
    class UrlHistory /* : IUrlHistoryStg[2] */ { }
    #endregion
}

