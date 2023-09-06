// Credit: ThiefMaster@GitHub, adrian[at]planetcoding[dot]net
// https://github.com/ThiefMaster/coreaudio-dotnet/blob/master/CoreAudio/CPolicyConfigClient.cs
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace AudioSwitch.COM
{
    [ComImport, Guid("294935CE-F637-4E7C-A41B-AB255460B862")]
    internal class _CPolicyConfigVistaClient
    {
    }

    [Guid("568b9108-44bf-40b4-9006-86afe5b5a620"),
InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IPolicyConfigVista
    {
        [PreserveSig]
        int GetMixFormat();
        [PreserveSig]
        int GetDeviceFormat();
        [PreserveSig]
        int SetDeviceFormat();
        [PreserveSig]
        int GetProcessingPeriod();
        [PreserveSig]
        int SetProcessingPeriod();
        [PreserveSig]
        int GetShareMode();
        [PreserveSig]
        int SetShareMode();
        [PreserveSig]
        int GetPropertyValue();
        [PreserveSig]
        int SetPropertyValue();
        [PreserveSig]
        int SetDefaultEndpoint([MarshalAs(UnmanagedType.LPWStr)] string wszDeviceId, ERole eRole);
        [PreserveSig]
        int SetEndpointVisibility();
    }

    public class CPolicyConfigVistaClient : IDisposable
    {
        private IPolicyConfigVista _policyConfigVistaClient = new _CPolicyConfigVistaClient() as IPolicyConfigVista;

        public void SetDefaultDevice(string deviceID)
        {
            _policyConfigVistaClient.SetDefaultEndpoint(deviceID, ERole.Console);
            _policyConfigVistaClient.SetDefaultEndpoint(deviceID, ERole.Multimedia);
            _policyConfigVistaClient.SetDefaultEndpoint(deviceID, ERole.Communications);
        }

        public void Dispose()
        {
            Marshal.ReleaseComObject(_policyConfigVistaClient);
        }
    }
}
