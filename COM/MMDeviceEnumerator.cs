// https://github.com/microsoft/psi
﻿// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

namespace AudioSwitch.COM // Microsoft.Psi.Audio.ComInterop
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// MMDeviceEnumerator COM class declaration.
    /// </summary>
    [ComImport]
    [Guid(Guids.MMDeviceEnumeratorCLSIDString)]
    internal class MMDeviceEnumerator
    {
    }
}
