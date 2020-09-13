using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml;
using Syncfusion.XlsIO;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        private async void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
           // var handler = new HttpClientHandler() { UseCookies = true };
           //  var client = new HttpClient(handler); 
           //  client.DefaultRequestHeaders.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.140 Safari/537.36 Edge/17.17134");
           //  client.DefaultRequestHeaders.Add("Referer", "http://10.232.0.7:8888/His/doctor/bus/search/patient.do");
           //
           //  client.DefaultRequestHeaders.Add("Cookie", "HisuserId=AA0" +
           //      "34DEB8967095F04C1F7CD9D69FCA4; HisuserName=011" +
           //      "9gk; Hisname=%E9%AB%98%E5%87%AF; HisorganizeCo" +
           //      "de=569823095; HisorganizeId=56072; Hiso" +
           //      "rganizeSid=0.10.11.2411.2412.56072.; Hisorg" +
           //      "anizeName=%E6%96%B0%E5%BB%BA%E8%B7%AF%E7%A4%BE%E5" +
           //      "%8C%BA%E5%8D%AB%E7%94%9F%E6%9C%8D%E5%8A%A1%E4%B8%AD%" +
           //      "E5%BF%83; HisareaId=2411; HisareaCode=410184; HisareaNam" +
           //      "e=%E6%96%B0%E9%83%91%E5%B8%82; HisareaSid=0.10.11.2411.; H" +
           //      "isoperId=DBF1D7C3D7C0E766415A231FFC9235F0; HisofficeId=47" +
           //      "392; validateCode=5EE6D2A8557AFD15DDB29371B4288A05; userId" +
           //      "=AA034DEB8967095F04C1F7CD9D69FCA4; userName=%E9%AB%98%E5%" +
           //      "87%AF; userCode=0119gk; organizeId=E4D984B9EBF94A19A516ADC" +
           //      "2583233F9; organizeName=%E6%96%B0%E5%BB%BA%E8%B7%AF%E7%" +
           //      "A4%BE%E5%8C%BA%E5%8D%AB%E7%94%9F%E6%9C%8D%E5%8A%A1%E4" +
           //      "%B8%AD%E5%BF%83; countyId=2411; countyName=%E6%96%B0%E" +
           //      "9%83%91%E5%B8%82; countyCode=410184; theme=green; home=" +
           //      "index3; activeRoleId=19068; __guid=83709462.314919" +
           //      "8815292528600.1599984462097.3162; monitor_count=1; " +
           //      "loginTime=1599984468649; JSESSIONID=61iGgcMVGwi7dHezqtzfr" +
           //      "uXHW8xDzlkyiHB_BbFaLPpVLIfBdK30!-637378030");
           //  List<KeyValuePair<string, string>> list = new List<KeyValuePair<string, string>>()
           //  {
           //      new KeyValuePair<string, string>("STARTDATE","2019-09-01"),
           //        new KeyValuePair<string, string>("ENDDATE","2019-12-30"),
           //
           //              new KeyValuePair<string, string>("NAME",""),
           //                new KeyValuePair<string, string>("CARDID",""),
           //                 new KeyValuePair<string, string>("GHDH",""),
           //                  new KeyValuePair<string, string>("ZSY",""),
           //                   new KeyValuePair<string, string>("YLZH",""),
           //                    new KeyValuePair<string, string>("OPERNAME",""),
           //                     new KeyValuePair<string, string>("thirtyFive","0"),
           //                      new KeyValuePair<string, string>("gaoxueya","0"),
           //                      new KeyValuePair<string, string>("chuanranbing","0"),
           //                      new KeyValuePair<string, string>("tangniaobing","0"),
           //                      new KeyValuePair<string, string>("source",""),
           //                      new KeyValuePair<string, string>("CLINRESUNAME",""),
           //                      new KeyValuePair<string, string>("officeid",""),
           //                       new KeyValuePair<string, string>("REGITRACK",""),
           //                       new KeyValuePair<string, string>("jzhzcxbbgs","0"),
           //
           //
           //  };
           //
           //  var s = await client.PostAsync("http://10.232.0.7:8888/His/search/importOut.do", new FormUrlEncodedContent(list));
           //  var st = await s.Content.ReadAsByteArrayAsync();
           //  File.WriteAllBytes("./2019Season3.xml", st);
           //  MessageBox.Show("完成");
            var text=  File.ReadAllText("2019Season1.xml", Encoding.UTF8);
            var text2=  File.ReadAllText("2019Season2.xml", Encoding.UTF8);
            var text3=  File.ReadAllText("2019Season3.xml", Encoding.UTF8);
            var text4 = text + text2 + text3;
            var collection= Regex.Matches(text4, @"\<CLINRESUNAME\>\<!\[CDATA\[(.*?)\]\]\>\<\/CLINRESUNAME\>");
            tb1.Text = collection.Count.ToString();
            var app = new ExcelEngine();
            var workbook= app.Excel.Workbooks.Create();
            var worksheet = workbook.Worksheets.Create("统计表");
            
            for (int i = 1; i < collection.Count; i++)
            {
                worksheet[i, 1].Value = collection[i].Groups[1].Value;
            }
            
            var steam=new FileStream("2019.xlsx",FileMode.Create);
           
            workbook.SaveAs(steam);
            MessageBox.Show("成功");
        }
    }
}