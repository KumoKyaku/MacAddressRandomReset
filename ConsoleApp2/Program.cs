using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Win32;
using System.Net.NetworkInformation;
using System.Management;
using System.Threading;
using System.Runtime.InteropServices;
using NETCONLib;
using DynamicMAC;
using System.Collections;

namespace ConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {
            Reset();
        }

        private static void Reset()
        {
            MACHelper mac = new MACHelper();
            int defaultIPAdpter = mac.GetIndex(out string defualtMac);

            if (string.IsNullOrEmpty(defualtMac))
            {
                Console.WriteLine("Mac is Null ! Press F1 to try again...");
                var key2 = Console.ReadKey();
                if (key2.Key == ConsoleKey.F1)
                {
                    Console.WriteLine("               ");
                    Reset();
                }
            }
            else
            {
                Console.WriteLine($"OldMac: {defualtMac}");
                mac.ResetMACAddress();
                Console.Write("Wait.");
                Thread.Sleep(1500);
                string newMac = mac.GetMac(defaultIPAdpter);

                while (string.IsNullOrEmpty(newMac))
                {
                    Console.Write(".");
                    newMac = mac.GetMac(defaultIPAdpter);
                    Thread.Sleep(500);
                }
                Console.WriteLine(".");
                Console.WriteLine($"NewMac: {newMac}");

                Console.WriteLine("Press F1 to ReSet Mac again!. Press any otherKey Exit...");
                var key = Console.ReadKey();
                if (key.Key == ConsoleKey.F1)
                {
                    Console.WriteLine("               ");
                    Reset();
                }
            }
        }
    }
}


namespace DynamicMAC
{
    public class MACHelper
    {
        [DllImport("wininet.dll")]
        private extern static bool InternetGetConnectedState(int Description, int ReservedValue);
        /// <summary>
        /// 是否能连接上Internet
        /// </summary>
        /// <returns></returns>
        public bool IsConnectedToInternet()
        {
            int Desc = 0;
            return InternetGetConnectedState(Desc, 0);
        }
        /// <summary>
        /// 获取MAC地址
        /// </summary>
        public string GetMACAddress()
        {
            //得到 MAC的注册表键
            RegistryKey macRegistry = Registry.LocalMachine.OpenSubKey("SYSTEM").OpenSubKey("CurrentControlSet").OpenSubKey("Control")
                .OpenSubKey("Class").OpenSubKey("{4D36E972-E325-11CE-BFC1-08002bE10318}");
            IList<string> list = macRegistry.GetSubKeyNames().ToList();
            IPGlobalProperties computerProperties = IPGlobalProperties.GetIPGlobalProperties();
            NetworkInterface[] nics = NetworkInterface.GetAllNetworkInterfaces();
            var adapter = nics.First(o => o.Name == "本地连接");
            if (adapter == null)
                return null;
            return string.Empty;
        }
        /// <summary>
        /// 设置MAC地址
        /// </summary>
        /// <param name="newMac"></param>
        public void SetMACAddress(string newMac)
        {
            string macAddress;
            string index = GetAdapterIndex(out macAddress);
            if (index == null)
                return;
            //得到 MAC的注册表键
            RegistryKey macRegistry = Registry.LocalMachine.OpenSubKey("SYSTEM").OpenSubKey("CurrentControlSet").OpenSubKey("Control")
                .OpenSubKey("Class").OpenSubKey("{4D36E972-E325-11CE-BFC1-08002bE10318}").OpenSubKey(index, true);
            if (string.IsNullOrEmpty(newMac))
            {
                newMac = CreateNewMacAddress();
            }

            macRegistry.SetValue("NetworkAddress", newMac, RegistryValueKind.String);

            macRegistry.OpenSubKey("Ndi", true).OpenSubKey("params", true).OpenSubKey("NetworkAddress", true).SetValue("default", newMac);
            macRegistry.OpenSubKey("Ndi", true).OpenSubKey("params", true).OpenSubKey("NetworkAddress", true).SetValue("ParamDesc", "Network Address");

            Thread oThread = new Thread(new ThreadStart(ReConnect));//new Thread to ReConnect
            oThread.Start();
        }
        /// <summary>
        /// 重设MAC地址
        /// </summary>
        public void ResetMACAddress()
        {
            SetMACAddress(string.Empty);
        }
        /// <summary>
        /// 重新连接
        /// </summary>
        private void ReConnect()
        {
            NetSharingManagerClass netSharingMgr = new NetSharingManagerClass();
            INetSharingEveryConnectionCollection connections = netSharingMgr.EnumEveryConnection;
            foreach (INetConnection connection in connections)
            {
                INetConnectionProps connProps = netSharingMgr.get_NetConnectionProps(connection);
                if (connProps.MediaType == tagNETCON_MEDIATYPE.NCM_LAN)
                {
                    connection.Disconnect(); //禁用网络
                    connection.Connect();    //启用网络
                }
            }
        }
        /// <summary>
        /// 生成随机MAC地址
        /// </summary>
        /// <returns></returns>
        public string CreateNewMacAddress()
        {
            //return "0016D3B5C493";
            int min = 0;
            int max = 16;
            Random ro = new Random();
            var sn = string.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}",
               ro.Next(min, max).ToString("x"),//0
               ro.Next(min, max).ToString("x"),//
               ro.Next(min, max).ToString("x"),
               ro.Next(min, max).ToString("x"),
               ro.Next(min, max).ToString("x"),
               ro.Next(min, max).ToString("x"),//5
               ro.Next(min, max).ToString("x"),
               ro.Next(min, max).ToString("x"),
               ro.Next(min, max).ToString("x"),
               ro.Next(min, max).ToString("x"),
               ro.Next(min, max).ToString("x"),//10
               ro.Next(min, max).ToString("x")
                ).ToUpper();
            return sn;
        }
        /// <summary>
        /// 得到Mac地址及注册表对应Index
        /// </summary>
        /// <param name="macAddress"></param>
        /// <returns></returns>
        public string GetAdapterIndex(out string macAddress)
        {
            int indexString = GetIndex(out macAddress);
            if (macAddress == string.Empty)
                return null;
            else
                return indexString.ToString().PadLeft(4, '0');
        }

        public int GetIndex(out string macAddress)
        {
            ManagementClass oMClass = new ManagementClass("Win32_NetworkAdapterConfiguration");
            ManagementObjectCollection colMObj = oMClass.GetInstances();
            macAddress = string.Empty;
            int indexString = 0;
            foreach (ManagementObject objMO in colMObj)
            {
                if (objMO["MacAddress"] != null && (bool)objMO["IPEnabled"] == true)
                {
                    macAddress = objMO["MacAddress"].ToString().Replace(":", "");
                    break;
                }
                indexString++;
            }
            return indexString;
        }

        public string GetMac(int index)
        {
            ManagementClass oMClass = new ManagementClass("Win32_NetworkAdapterConfiguration");
            ManagementObjectCollection colMObj = oMClass.GetInstances();
            string macAddress = string.Empty;
            int indexString = 0;
            
            foreach (ManagementObject objMO in colMObj)
            {
                if (indexString == index)
                {
                    if (objMO["MacAddress"] != null && (bool)objMO["IPEnabled"] == true)
                    {
                        macAddress = objMO["MacAddress"].ToString().Replace(":", "");
                        break;
                    }
                }

                indexString++;
            }
            return macAddress;
        }
    }
}