using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace NIM.Tester
{
    [TestClass]
    public class TcpClientTest
    {

        [TestMethod]
        public void Connect()
        {

            using (var tcpClient = new TcpClient())
            {
                tcpClient.Connect(IPAddress.Parse("192.168.1.110"), 10086);

                 
            }

        }
    }
}
