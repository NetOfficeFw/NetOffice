using System;
using System.Collections.Generic;
using System.Threading;
using System.Text;

namespace NOToolsTests.ConsoleMonitor1
{
    class Program
    {
        static void Main(string[] args)
        {
            int firstCounter = 10;
            int secondCounter = 10;
            int thirdCounter = 10;

            Console.WriteLine("NOToolsTests.ConsoleMonitor1{0}Press key to start...", Environment.NewLine);
            Console.ReadKey();
            Console.WriteLine("Running");

            // hierarchy messages to the main console
            string log0 = NetOffice.DebugConsole.SendPipeConsoleMessage(null, String.Format("Level0 Start"));
            for (int i = 1; i <= firstCounter; i++)
            {
                string log1 = NetOffice.DebugConsole.SendPipeConsoleMessage(null, String.Format("Level1 {0}", i), log0);
                for (int y = 0; y < secondCounter; y++)
                {
                    string log2 = NetOffice.DebugConsole.SendPipeConsoleMessage(null, String.Format("Level2 {0}", y), log1);
                    for (int z = 0; z < thirdCounter; z++)
                        NetOffice.DebugConsole.SendPipeConsoleMessage(null, String.Format("Level3 {0}", z), log2);
                }
            }
            NetOffice.DebugConsole.SendPipeConsoleMessage(null, String.Format("Level0 End"));

            // some messages to custom consoles
            for (int i = 0; i < firstCounter; i++)
                NetOffice.DebugConsole.SendPipeConsoleMessage("CustomConsole1", String.Format("Message {0}", i));
            for (int i = 0; i < firstCounter; i++)
                NetOffice.DebugConsole.SendPipeConsoleMessage("CustomConsole2", String.Format("Message {0}", i));

            // channel messages        
            NetOffice.DebugConsole.SendPipeChannelMessage("Channel 1", "Message1-1");
            NetOffice.DebugConsole.SendPipeChannelMessage("Channel 2", "Message2-1");
            NetOffice.DebugConsole.SendPipeChannelMessage("Channel 3", "Message3-1");
            NetOffice.DebugConsole.SendPipeChannelMessage("Channel 4", "Message4-1");
            NetOffice.DebugConsole.SendPipeChannelMessage("Channel 5", "Message5-1");
            NetOffice.DebugConsole.SendPipeChannelMessage("Channel 6", "Message6-1");
            NetOffice.DebugConsole.SendPipeChannelMessage("Channel 7", "Message7-1");
            NetOffice.DebugConsole.SendPipeChannelMessage("Channel 2", "Message2-2");
            NetOffice.DebugConsole.SendPipeChannelMessage("Channel 2", "Message2-3");
            NetOffice.DebugConsole.SendPipeChannelMessage("Channel 2", "Message2-4");
            
            Console.WriteLine("Press key to exit...");
            Console.Read();

            return;
        }
    }
}
