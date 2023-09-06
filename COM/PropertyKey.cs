// https://github.com/microsoft/psi
﻿// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

namespace AudioSwitch.COM // Microsoft.Psi.Audio.ComInterop
{
    using System;

    /// <summary>
    /// PROPERTYKEY struct.
    /// </summary>
    internal struct PropertyKey
    {
        /// <summary>
        /// Format ID.
        /// </summary>
        internal Guid FormatId;

        /// <summary>
        /// Property ID.
        /// </summary>
        internal int PropertyId;
    }
}