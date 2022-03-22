using System;
using static System.Console;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security.Cryptography;
using System.Web;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Configuration;
using System.Net;
using System.IO;
using System.Threading;
using System.Data;
using System.Xml;
using System.Collections;
using System.Windows;

namespace Aliyun_Robot_test
{
    internal class Program
    {
        //0=加签，1=关键词,2=加签+关键词,3=IP地址(暂不支持)
        public static int ori_mode = 0;

        private static string ori_url = "https://oapi.dingtalk.com/robot/send";
        //亿雨
        private static string ori_access_token = "f1c99c1939efddfc0d04528c3a63d48da588380233de653893cbfda611801d15";
        //亿雨
        private static string ori_secret_token = "SEC4493b7d3c9e391556f7f13e811f270c70acca4c2245e48c42c7d66ace8bc8e1a";

        public static string ori_keyword = null;

        public static string config_path = "conf.json";

        //瑞恩
        //private static string ori_access_token = "3e990a4842062ecaa9c255a2daf15f56eb9dec99a7dfa9b5d0ddb82ea37e2a6f";

        //读取列表
        public class config 
        {
            public int mode { set; get; } = 0;

            public string url { set; get; } = ori_url;
            public string access_token { set; get; } = null;

            public string secret_token { set; get; } = null;

            public string keyword { set; get; } = null;

            
        }

        //https://blog.csdn.net/weixin_34247299/article/details/93766897
        static bool OpenCSVFile(ref DataTable mycsvdt, string filepath)
        {
            string strpath = filepath; //csv文件的路径
            try
            {
                int intColCount = 0;
                bool blnFlag = true;

                DataColumn mydc;
                DataRow mydr;

                string strline;
                string[] aryline;
                StreamReader mysr = new StreamReader(strpath, System.Text.Encoding.Default);

                while ((strline = mysr.ReadLine()) != null)
                {
                    aryline = strline.Split(new char[] { ',' });

                    //给datatable加上列名
                    if (blnFlag)
                    {
                        blnFlag = false;
                        intColCount = aryline.Length;
                        int col = 0;
                        for (int i = 0; i < aryline.Length; i++)
                        {
                            col = i + 1;
                            mydc = new DataColumn(col.ToString());
                            mycsvdt.Columns.Add(mydc);
                        }
                    }

                    //填充数据并加入到datatable中
                    mydr = mycsvdt.NewRow();
                    for (int i = 0; i < intColCount; i++)
                    {
                        mydr[i] = aryline[i];
                    }
                    mycsvdt.Rows.Add(mydr);
                }
                return true;

            }
            catch (Exception e)
            {
                WriteLine(e.Message);
                return false;
            }
        }
        static string get_time_stamp(long timestamp) 
        {
            string stringToSign = timestamp + "\n" + ori_secret_token;
            var encoding = new System.Text.ASCIIEncoding();
            byte[] keyByte = encoding.GetBytes(ori_secret_token);
            byte[] messageBytes = encoding.GetBytes(stringToSign);
            using (var hmacsha256 = new HMACSHA256(keyByte))
            {
                byte[] hashmessage = hmacsha256.ComputeHash(messageBytes);
                return HttpUtility.UrlEncode(Convert.ToBase64String(hashmessage), Encoding.UTF8);
            }
        }
        /// <summary>
        /// 发送GET请求
        /// </summary>
        /// <param name="url">请求URL，如果需要传参，在URL末尾加上“？+参数名=参数值”即可</param>
        /// <returns></returns> 
        static string HttpGet(string url)
        {
            //创建
            HttpWebRequest httpWebRequest = (HttpWebRequest)HttpWebRequest.Create(url);
            //设置请求方法
            httpWebRequest.Method = "GET";
            //请求超时时间
            httpWebRequest.Timeout = 20000;
            //发送请求
            HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            //利用Stream流读取返回数据
            StreamReader streamReader = new StreamReader(httpWebResponse.GetResponseStream(), Encoding.UTF8);
            //获得最终数据，一般是json
            string responseContent = streamReader.ReadToEnd();

            streamReader.Close();
            httpWebResponse.Close();

            return responseContent;
        }
        /// <summary>
        /// 发送POST请求
        /// </summary>
        /// <param name="url">请求URL</param>
        /// <param name="data">请求参数</param>
        /// <returns></returns>
        /// 
        static string HttpPost(string url, string data)
        {
            HttpWebRequest httpWebRequest = (HttpWebRequest)HttpWebRequest.Create(url);

            //字符串转换为字节码
            byte[] bs = Encoding.UTF8.GetBytes(data);
            //参数类型，这里是json类型
            //还有别的类型如"application/x-www-form-urlencoded"，不过我没用过(逃
            httpWebRequest.ContentType = "application/json";
            //参数数据长度
            httpWebRequest.ContentLength = bs.Length;
            //设置请求类型
            httpWebRequest.Method = "POST";
            //设置超时时间
            httpWebRequest.Timeout = 20000;
            //将参数写入请求地址中
            httpWebRequest.GetRequestStream().Write(bs, 0, bs.Length);

            //发送请求
            HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            //读取返回数据
            StreamReader streamReader = new StreamReader(httpWebResponse.GetResponseStream(), Encoding.UTF8);
            string responseContent = streamReader.ReadToEnd();

            streamReader.Close();
            httpWebResponse.Close();
            httpWebRequest.Abort();

            return responseContent;
        }
        static void SendText(string Content) 
        {
            string MessageUrl = null;
            if (ori_mode == 0)
            {
                MessageUrl =
                ori_url + "?access_token=" +
                ori_access_token + "&timestamp=" +
                (long)(DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds +
                "&sign=" + get_time_stamp((long)(DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds);
            }
            if (ori_mode == 1)
            {
                MessageUrl =
                ori_url + "?access_token=" +
                ori_access_token;
                Content = ori_keyword + ":" + Content;
            }
            if (ori_mode == 2)
            {
                MessageUrl =
                ori_url + "?access_token=" +
                ori_access_token + "&timestamp=" +
                (long)(DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds +
                "&sign=" + get_time_stamp((long)(DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds);
                Content = ori_keyword + ":" + Content;
            }
            var json_req = new
            {
                msgtype = "text", //推送类型
                text = new
                {
                    content = Content
                }
            };
            WriteLine(Content);
            string jsonRequest = JsonConvert.SerializeObject(json_req);//将对象转换为json
            WriteLine(MessageUrl);
            string rep = HttpPost(MessageUrl, jsonRequest);
            WriteLine(rep);

        }

        static int NewConfigJson() 
        {
            config conf = new config()
            {
                mode = 0,
                url = ori_url,
                access_token = null,
                secret_token = null
            };
            string confJson = JsonConvert.SerializeObject(conf);
            using (StreamWriter sw = new StreamWriter(path: config_path))
            {
                sw.WriteLine(confJson);
                return 0;
            }
            //return -1;
        }

        static int WriteConfigJson(config conf)
        {
            string confJson = JsonConvert.SerializeObject(conf);
            using (StreamWriter sw = new StreamWriter(path: config_path))
            {
                sw.WriteLine(confJson);
                return 0;
            }
            //return -1;
        }

        static JObject GetJsonContects(string jsonfile)
        {
            if (!File.Exists(jsonfile)) 
            {
                WriteLine("无法找到" + jsonfile + "文件！");
                WriteLine("按回车返回！");
                ReadLine();
                return null;
            }
            JObject job = new JObject();
            using (System.IO.StreamReader file = System.IO.File.OpenText(jsonfile))
            {
                using (JsonTextReader reader = new JsonTextReader(file))
                {
                    job = (JObject)JToken.ReadFrom(reader);
                }
            }
            return job;
        }

        static void ReadConfigJson()
        {
            JObject job = GetJsonContects(config_path);
            ori_mode = int.Parse(job["mode"].ToString()); 
            ori_url = job["url"].ToString();
            ori_access_token = job["access_token"].ToString();
            ori_secret_token = job["secret_token"].ToString();
            ori_keyword = job["keyword"].ToString();
        }

        static string[] GetWeacher() 
        {
            List<string> Weachers = new List<string>();
            string MessageUrl =
                "https://tianqiapi.com/free/day" + "?appid=" +
                "54172754" + "&appsecret=" +
                "OwqCi0x9" +
                "&city=" + "浦东新区";
            string responseContent = HttpGet(MessageUrl);
            JObject job = JObject.Parse(responseContent);
            Thread.Sleep(735);
            WriteLine("天气获取成功！正在转换数据");
            Weachers.Add(job["city"].ToString()); //城市
            WriteLine("区域:" + job["city"]);
            Weachers.Add(job["wea"].ToString());  //天气
            WriteLine("天气:" + job["wea"]);
            Weachers.Add(job["tem"].ToString() + "°");  //平均温度
            WriteLine("平均温度:" + job["tem"] + "°");
            Thread.Sleep(735);
            Weachers.Add(job["tem_day"].ToString() + "°");  //最大温度
            Weachers.Add(job["tem_night"].ToString() + "°");  //最低温度
            Weachers.Add(job["win"].ToString());  //风向
            Weachers.Add(job["win_meter"].ToString());  //风速
            string[] rsp =  Weachers.ToArray();
            Thread.Sleep(735);
            //Console.WriteLine(rb);
            return rsp;
        }
        static void SendWeacher(string[] WEA) 
        {
            string msg = "正在获取天气数据...";
            SendText(msg);
            Thread.Sleep(1000);
            string msg2 =
                " 城市:" +
                WEA[0] +
                " 天气:" +
                WEA[1] +
                " 平均温度:" +
                WEA[2];
            string msg3 =
                " 最大温度 " +
                WEA[3] +
                " 最低温度 " +
                WEA[4] +
                " 风向 " +
                WEA[5] +
                "_" +
                WEA[6] +
                "";
            SendText(msg2);
            Thread.Sleep(735);
            SendText(msg3);
        }

        static string getDayofWeek(string week)
        {
            switch (week) 
            {
                case "周一":
                    return "周二";
                case "周二":
                    return "周三";
                case "周三":
                    return "周四";
                case "周四":
                    return "周五";
                default:
                    return "周一";
            }
        }
        static int GetClassofWeek(string classes)
        {
            switch (classes) 
            {
                case "休息+室内锻炼活动":
                    return -1;
                case "休息+眼保健操":
                    return -1;
                case "午餐+休息+自主阅读":
                    return 1;
                case "":
                    return 1;
                default:
                    return 0;
            }
        }
        static string[] GetClassNext(string path) 
        {
            //行
            List<DataRow> Ldata = new List<DataRow>();
            //列
            List<DataColumn> Rdata = new List<DataColumn>();
            List<DateTime> times_up = new  List<DateTime>();
            List<DateTime> times_down = new List<DateTime>();
            List<string> Day_class = new List<string>();
            DataTable dataT = new DataTable();
            OpenCSVFile(ref dataT, path);
            for (var i = 0; i < dataT.Rows.Count; i++) 
            {
                //每一行
                DataRow dataRow = dataT.Rows[i];
                Ldata.Add(dataRow); 
            }
            for (var i = 0; i < dataT.Columns.Count; i++)
            {
                //每一列
                DataColumn dataColumns = dataT.Columns[i];
                Rdata.Add(dataColumns);
            }
            //第二行
            if (Ldata[0].ItemArray.Contains("时间"))
            {
                //遍历每一列，第一列除外
                for (var i = 1; i < Ldata.Count; i++)
                {
                    //遍历每一列的数据
                    for (var j = 0; j < Ldata[i].ItemArray.Length; j++) 
                    {
                        //获取时间
                        if (j == 1) 
                        {
                            var _str = Ldata[i].ItemArray[j].ToString();
                            var _str_spl = _str.Split('-');
                            times_up.Add(DateTime.Parse(_str_spl[0]));
                            //WriteLine(_str_spl[0]);
                            times_down.Add(DateTime.Parse(_str_spl[1]));
                            //WriteLine(_str_spl[1]);
                        }
                        //获取课程(根据时间)
                        var _time = DateTime.Now.ToString("ddd");
                        _time = getDayofWeek(_time);
                        Day_class.Add(getDayofWeek(DateTime.Now.ToString("ddd")) + ":");
                        //WriteLine("i = " + i);
                        //WriteLine("j = " + j);
                        for (var k = 2; k <Ldata[0].ItemArray.Length; k++) 
                        {
                            
                            WriteLine(_time +"---------------------"+ (string)Ldata[0].ItemArray[k]);

                            if ((string)Ldata[0].ItemArray[k] == _time) 
                            {
                                for (var m = 2; m < dataT.Rows.Count; m++) 
                                {
                                    WriteLine("---------------------");
                                    //WriteLine("m = " + m);
                                    //WriteLine("k = " + k);
                                    //WriteLine("dataT.Rows.Count = " + dataT.Rows.Count);
                                    if (GetClassofWeek((string)dataT.Rows[m][k]) == 0)
                                    {
                                        Day_class.Add((string)dataT.Rows[m][k]);
                                        WriteLine((string)dataT.Rows[m][k]);
                                    }
                                    else if (GetClassofWeek((string)dataT.Rows[m][2]) == 1) 
                                    {
                                        Day_class.Add("|");
                                        WriteLine("|");
                                    }
                                    
                                }
                                /*
                                for (int m = 0; m < Day_class.Count; m++)
                                {
                                    if (DateTime.Now.ToString("t").CompareTo(times_up[m]) >0)
                                    {
                                        WriteLine("现在是上课时间");
                                    }
                                    if (DateTime.Now.ToString("t").CompareTo(times_down[m]) > 0)
                                    {
                                        WriteLine("现在是下课时间");
                                    }
                                }
                                */
                                WriteLine("返回值:" + Day_class.ToArray());
                                return Day_class.ToArray();
                            }
                        }
                    }
                }
            }
            //Write(Ldata.Count);
            return new string[0];
        }

        static void SendClassNext(string[] classes) 
        {
            string _str = null;
            foreach (var c in classes) 
            {
                _str = _str +" "+ c;
            }
            SendText("正在获取明日课程...");
            Thread.Sleep(735);
            SendText("明日课程列表:");
            SendText(_str);
        }

        static string GetAccessToken(string url) 
        {
            url=url.Substring(url.IndexOf('=')+1);
            WriteLine("access_token获取成功！");
            WriteLine("access_token:" + url);
            return url;
        }
        static void Main(string[] args)
        {
            if (!File.Exists(config_path))
            {
                config conf = new config();
                WriteLine("找不到机器人配置JSON，程序初始化");
                WriteLine("正在初始化配置文件...");
                NewConfigJson();
                Clear();
                WriteLine("请选择当前模式");
                WriteLine("0.加签模式 1.关键词模式 2.加签+关键词");
                var _value = ReadLine();
                conf.mode = int.Parse(_value);
                Thread.Sleep(200);
                WriteLine("---------------------------");
                WriteLine("请复制粘贴webhook");
                _value = ReadLine();
                conf.access_token = GetAccessToken(_value);
                Thread.Sleep(200);
                WriteLine("---------------------------");
                if (conf.mode == 0 || conf.mode==2)
                {
                    WriteLine("请复制粘贴密钥");
                    _value = ReadLine();
                    conf.secret_token = _value;

                }
                if (conf.mode == 1 || conf.mode == 2)
                {
                    WriteLine("请复制粘贴关键词");
                    _value = ReadLine();
                    conf.keyword = _value;
                }

                WriteConfigJson(conf);
            }
            ReadConfigJson();
            Console.WriteLine("正在启动机器人...");
            Thread.Sleep(735);
            while (true) 
            {
                Clear();
                WriteLine("选项:1.天气 2.明日课程表提醒 N.自定义输入 D.删除配置文件并重载");
                string Input = ReadLine();
                if (Input == "1") 
                {
                    var wea = GetWeacher();
                    ReadLine();
                    Clear();
                    SendWeacher(wea);
                    ReadLine();
                }
                if (Input == "2") 
                {
                    var cla = GetClassNext("kechengbiao.csv");
                    ReadLine();
                    Clear();
                    SendClassNext(cla);
                    ReadLine();
                }
                if(Input == "N") 
                {
                    WriteLine("请输入想要发送的消息");
                    string msg = Console.ReadLine();
                    SendText(msg);
                    ReadLine();
                }
                if (Input == "D")
                {
                    File.Delete(config_path);
                    System.Windows.Forms.Application.Restart();
                }
            }
        }
    }
}
