// 
// Copyright (c) Microsoft Corporation. All rights reserved.
// 
// Microsoft Public License (MS-PL)
// 
// This license governs use of the accompanying software. If you use the
// software, you accept this license. If you do not accept the license, do not
// use the software.
// 
// 1. Definitions
// 
//   The terms "reproduce," "reproduction," "derivative works," and
//   "distribution" have the same meaning here as under U.S. copyright law. A
//   "contribution" is the original software, or any additions or changes to
//   the software. A "contributor" is any person that distributes its
//   contribution under this license. "Licensed patents" are a contributor's
//   patent claims that read directly on its contribution.
// 
// 2. Grant of Rights
// 
//   (A) Copyright Grant- Subject to the terms of this license, including the
//       license conditions and limitations in section 3, each contributor
//       grants you a non-exclusive, worldwide, royalty-free copyright license
//       to reproduce its contribution, prepare derivative works of its
//       contribution, and distribute its contribution or any derivative works
//       that you create.
// 
//   (B) Patent Grant- Subject to the terms of this license, including the
//       license conditions and limitations in section 3, each contributor
//       grants you a non-exclusive, worldwide, royalty-free license under its
//       licensed patents to make, have made, use, sell, offer for sale,
//       import, and/or otherwise dispose of its contribution in the software
//       or derivative works of the contribution in the software.
// 
// 3. Conditions and Limitations
// 
//   (A) No Trademark License- This license does not grant you rights to use
//       any contributors' name, logo, or trademarks.
// 
//   (B) If you bring a patent claim against any contributor over patents that
//       you claim are infringed by the software, your patent license from such
//       contributor to the software ends automatically.
// 
//   (C) If you distribute any portion of the software, you must retain all
//       copyright, patent, trademark, and attribution notices that are present
//       in the software.
// 
//   (D) If you distribute any portion of the software in source code form, you
//       may do so only under this license by including a complete copy of this
//       license with your distribution. If you distribute any portion of the
//       software in compiled or object code form, you may only do so under a
//       license that complies with this license.
// 
//   (E) The software is licensed "as-is." You bear the risk of using it. The
//       contributors give no express warranties, guarantees or conditions. You
//       may have additional consumer rights under your local laws which this
//       license cannot change. To the extent permitted under your local laws,
//       the contributors exclude the implied warranties of merchantability,
//       fitness for a particular purpose and non-infringement.
//       

using System;
using System.Dynamic;

namespace Microsoft.ClearScript
{
    /// <summary>
    /// Defines options for exposing host resources to script code.
    /// </summary>
    [Flags]
    public enum HostItemFlags
    {
        /// <summary>
        /// Specifies that no options are selected.
        /// </summary>
        None = 0,

        /// <summary>
        /// Specifies that the host resource's members are to be exposed as global items in the
        /// script engine's root namespace.
        /// </summary>
        GlobalMembers = 0x00000001,

        /// <summary>
        /// Specifies that the host resource's non-public members are to be exposed.
        /// </summary>
        PrivateAccess = 0x00000002,

        /// <summary>
        /// Specifies that the host resource's dynamic members are not to be exposed. This option
        /// applies only to objects that implement <see cref="IDynamicMetaObjectProvider"/>.
        /// </summary>
        HideDynamicMembers = 0x00000004,

        /// <summary>
        /// Specifies that the script engine is to be given direct access to the exposed object if
        /// possible. This option, when supported, suppresses marshaling and hands off the object
        /// for script access without the host's involvement. It is currently supported only for
        /// COM objects exposed in Windows Script engines.
        /// </summary>
        DirectAccess = 0x00000008
    }
}
