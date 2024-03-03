using System;
using System.Runtime.InteropServices;
using NetOffice.Tools;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Tools;
using NetOffice.ExcelApi.Tools.Attributes;
/*
    Sample RTD Component
*/
namespace Excel05AddinCS4
{
    [COMRtdServer("A sample rtd server", 1)]
    [ProgId("MyRtdServerCS4.Server"), Guid("B114AC9F-9CC9-4B5A-89BA-A1BC29D5E963"), Codebase, Programmable]
    public class Server : RealtimeDataServer
    {
        private int TopicID { get; set; }

        protected override object ConnectData(int topicID, object strings, bool getNewValues)
        {
            TopicID = topicID;
            return GetTime();
        }

        protected override object RefreshData(int topicCount)
        {
            object[,] data = new object[2, 1];
            data[0, 0] = TopicID;
            data[1, 0] = GetTime();
            topicCount = 1;
            return data;
        }

        protected override void ServerTerminate()
        {

        }

        private string GetTime()
        {
            return "GetTime " + DateTime.Now.ToString();
        }
    }
}