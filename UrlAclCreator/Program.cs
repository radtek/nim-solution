using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Text;

namespace UrlAclCreator
{
    class Program
    {
        private const int PORT = 10086;
        static void Main(string[] args)
        {
            var ips = _GetIPs();
            var commandTexts = BulidCommand(ips);
            commandTexts.ForEach(commandText =>
            {
                RunCommand(commandText);
            });

            //let's double check

            Console.WriteLine("运行成功.");

            Console.ReadKey();
        }

        private static List<string> BulidCommand(List<string> ips)
        {
            var values = new List<string>();
            ips.ForEach(t =>
            {
                var commandText = $"netsh http add urlacl url=http://{t}:{PORT}/david/ user=everyone";
                values.Add(commandText);
            });
            return values;
            /*当前电脑名：static System.Environment.MachineName
            当前电脑所属网域：static System.Environment.UserDomainName
            当前电脑用户：static System.Environment.UserName
        https://www.cnblogs.com/babycool/p/3569183.html

            http://blog.csdn.net/huwei2003/article/details/24235367
            */
            //netsh http delete urlacl url=http://+:10086/david/


        }
        private static void RunCommand(string command)
        {
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            p.StartInfo.FileName = "cmd.exe";
            p.StartInfo.UseShellExecute = false;    //是否使用操作系统shell启动
            p.StartInfo.RedirectStandardInput = true;//接受来自调用程序的输入信息
            p.StartInfo.RedirectStandardOutput = true;//由调用程序获取输出信息
            p.StartInfo.RedirectStandardError = true;//重定向标准错误输出
            p.StartInfo.CreateNoWindow = true;//不显示程序窗口
            p.Start();//启动程序

            //向cmd窗口发送输入信息
            p.StandardInput.WriteLine(command + "&exit");

            p.StandardInput.AutoFlush = true;
            //p.StandardInput.WriteLine("exit");
            //向标准输入写入要执行的命令。这里使用&是批处理命令的符号，表示前面一个命令不管是否执行成功都执行后面(exit)命令，如果不执行exit命令，后面调用ReadToEnd()方法会假死
            //同类的符号还有&&和||前者表示必须前一个命令执行成功才会执行后面的命令，后者表示必须前一个命令执行失败才会执行后面的命令



            //获取cmd窗口的输出信息
            string output = p.StandardOutput.ReadToEnd();

            //StreamReader reader = p.StandardOutput;
            //string line=reader.ReadLine();
            //while (!reader.EndOfStream)
            //{
            //    str += line + "  ";
            //    line = reader.ReadLine();
            //}

            p.WaitForExit();//等待程序执行完退出进程
            p.Close();



        }


        private static List<string> _GetIPs()
        {
            return new List<string> { "+" };
            var hostName = Dns.GetHostName();//本机名  
                                             //System.Net.IPAddress[] addressList = Dns.GetHostByName(hostName).AddressList;//会警告GetHostByName()已过期，我运行时且只返回了一个IPv4的地址  
            var addressList = Dns.GetHostAddresses(hostName).ToList();//会返回所有地址，包括IPv4和IPv6  
            var ips = new List<string>();
            for (var i = 0; i < addressList.Count; i++)
            {
                var ipAddress = addressList[i];
                var str = ipAddress.ToString();

                if (str.Contains('.'))
                {
                    ips.Add(str);
                }
            }

            return ips;


        }
    }
}
