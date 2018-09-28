using FastStream.Log;
using FastStream.socket;
using System;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            // FlashLogger.Debug("ddd");
            NettyConfig.Port = 6666;
            
             NettyClient  client= new NettyClient();
            client.Connect();
            while(true)
            {
                client.SendData(NettyConfig.NettyEncod.GetBytes(DateTime.Now.ToString()));
            }

        }
    }
}
