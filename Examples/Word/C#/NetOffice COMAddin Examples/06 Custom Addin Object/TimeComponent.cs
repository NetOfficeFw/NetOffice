using System;
using System.Runtime.InteropServices;

namespace Word06AddinCS4
{
    [ComVisible(true), InterfaceType(ComInterfaceType.InterfaceIsDual), Guid("ED27733D-DA7E-4674-804A-D750A3737BD3")]
    public interface ITimeComponent
    {
        [DispId(1)]
        string GetTime();
    }

    public class TimeComponent : ITimeComponent
    {
        public string GetTime()
        {
            return DateTime.Now.ToString();
        }
    }
}