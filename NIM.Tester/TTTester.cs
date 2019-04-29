using DocumentFormat.OpenXml.Packaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NIM.CertificationGenerator.RadiationThermomater;
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
    public class TTTester
    {

        [TestMethod]
        public void TestReadFile()
        {
            var filePath = @"d:\Book1.xlsx";

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, true))
            {
                var i = NIM.Utilty.ExcelHelper.GetHasGoodFgColor(document, "Sheet1", "A1");

                i = NIM.Utilty.ExcelHelper.GetHasGoodFgColor(document, "Sheet1", "A2");
                i = NIM.Utilty.ExcelHelper.GetHasGoodFgColor(document, "Sheet1", "A3");


                //var wordPath = $@"C:\Users\Exten\Dropbox\nim\documents\templateconfs\rr-RGrr20172.docx";

                //new WordGenerator().Process(result, wordPath);
            }
        }


        [TestMethod]
        public void TCPTest()
        {
            var ip = "192.168.31.109";
            var port = 10086;
            var ipAddress = IPAddress.Parse(ip);
            var remoteEP = new IPEndPoint(ipAddress, port);
            using (var tcpClient = new TcpClient())
            {
                tcpClient.Connect(remoteEP);

                var requstString = "{'RequestType':0,'Playload':null}";
              

                var array = Encoding.Unicode.GetBytes(requstString);
                tcpClient.Client.Send(array);


                var recvBufs = new byte[1024 * 10 * 10];
                var recvByteLength = tcpClient.Client.Receive(recvBufs);

                var responseString = Encoding.Unicode.GetString(recvBufs, 0, recvByteLength);

               
            }
        }
    }
}
