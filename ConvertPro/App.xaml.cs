using System.Windows;

namespace ConvertPro
{
    public partial class App : Application
    {
        // WebView2 需要在 STA 线程上运行，WPF 默认就是 STA，无需额外配置
    }
}
