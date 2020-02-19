// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System.Collections.Generic;
using System.Data.Common;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace System.Data.OleDb
{
    internal sealed class OleDbPropertyInfo
    {
        public Guid _propertySet;
        public int _propertyID;
        public string _description;
        public string _lowercase;
        public Type _type;

        public int _flags;
        public int _vtype;
        public object _supportedValues;

        public object _defaultValue;
    }

    internal sealed class PropertyInfoSet : SafeHandle
    {
        private readonly int setCount;
        private IntPtr descBuffer;

        internal PropertyInfoSet(UnsafeNativeMethods.IDBProperties idbProperties, PropertyIDSet propIDSet) : base(IntPtr.Zero, true)
        {
            OleDbHResult hr;
            int propIDSetCount = propIDSet.Count;
            RuntimeHelpers.PrepareConstrainedRegions();
            try
            { }
            finally
            {
                hr = idbProperties.GetPropertyInfo(propIDSetCount, propIDSet, out this.setCount, out base.handle, out this.descBuffer);
            }
            if ((0 <= hr) && (ADP.PtrZero != handle))
            {
                SafeNativeMethods.Wrapper.ClearErrorInfo();
            }
        }

        public override bool IsInvalid
        {
            get
            {
                return ((IntPtr.Zero == base.handle) && (IntPtr.Zero == this.descBuffer));
            }
        }

        internal Dictionary<string, OleDbPropertyInfo> GetValues()
        {
            Dictionary<string, OleDbPropertyInfo> propertyLookup = null;

            bool mustRelease = false;
            RuntimeHelpers.PrepareConstrainedRegions();
            try
            {
                DangerousAddRef(ref mustRelease);
                if (ADP.PtrZero != this.handle)
                {
                    propertyLookup = new Dictionary<string, OleDbPropertyInfo>(StringComparer.OrdinalIgnoreCase);

                    IntPtr setPtr = this.handle;

                    for (int i = 0; i < setCount; ++i, setPtr = ADP.IntPtrOffset(setPtr, ODB.SizeOf_tagDBPROPINFOSET))
                    {
                        Tuple<int, IntPtr, Guid> propertyInfoSet = GetPropertyInfoSet(setPtr);
                        int infoCount = propertyInfoSet.Item1;
                        IntPtr infoPtr = propertyInfoSet.Item2;
                        for (int k = 0; k < infoCount; ++k, infoPtr = ADP.IntPtrOffset(infoPtr, ODB.SizeOf_tagDBPROPINFO))
                        {
                            OleDbPropertyInfo propertyInfo = RetrievePropertyInfo(infoPtr, propertyInfoSet.Item3);

                            propertyLookup[propertyInfo._lowercase] = propertyInfo;
                        }
                    }
                }
            }
            finally
            {
                if (mustRelease)
                {
                    DangerousRelease();
                }
            }
            return propertyLookup;
        }

        private OleDbPropertyInfo RetrievePropertyInfo(IntPtr infoPtr, Guid guidPropertySet)
        {
            OleDbPropertyInfo propertyInfo = new OleDbPropertyInfo();
            if (ODB.IsRunningOnX86)
            {
                tagDBPROPINFO_x86 propinfo = new tagDBPROPINFO_x86();
                Marshal.PtrToStructure(infoPtr, propinfo);
                propertyInfo._propertySet = guidPropertySet;
                propertyInfo._propertyID = propinfo.dwPropertyID;
                propertyInfo._flags = propinfo.dwFlags;
                propertyInfo._vtype = propinfo.vtType;
                propertyInfo._supportedValues = propinfo.vValue;
                propertyInfo._description = propinfo.pwszDescription;
                propertyInfo._lowercase = propinfo.pwszDescription.ToLowerInvariant();
                propertyInfo._type = PropertyInfoSet.FromVtType(propinfo.vtType);
            }
            else
            {
                tagDBPROPINFO propinfo = new tagDBPROPINFO();
                Marshal.PtrToStructure(infoPtr, propinfo);
                propertyInfo._propertySet = guidPropertySet;
                propertyInfo._propertyID = propinfo.dwPropertyID;
                propertyInfo._flags = propinfo.dwFlags;
                propertyInfo._vtype = propinfo.vtType;
                propertyInfo._supportedValues = propinfo.vValue;
                propertyInfo._description = propinfo.pwszDescription;
                propertyInfo._lowercase = propinfo.pwszDescription.ToLowerInvariant();
                propertyInfo._type = PropertyInfoSet.FromVtType(propinfo.vtType);
            }
            return propertyInfo;
        }

        private Tuple<int, IntPtr, Guid> GetPropertyInfoSet(IntPtr setPtr)
        {
            int infoCount;
            IntPtr infoPtr;
            Guid guidPropertySet;
            if (ODB.IsRunningOnX86)
            {
                tagDBPROPINFOSET_x86 propinfoset = new tagDBPROPINFOSET_x86();
                Marshal.PtrToStructure(setPtr, propinfoset);
                infoCount = propinfoset.cPropertyInfos;
                infoPtr = propinfoset.rgPropertyInfos;
                guidPropertySet = propinfoset.guidPropertySet;
            }
            else
            {
                tagDBPROPINFOSET propinfoset = new tagDBPROPINFOSET();
                Marshal.PtrToStructure(setPtr, propinfoset);
                infoCount = propinfoset.cPropertyInfos;
                infoPtr = propinfoset.rgPropertyInfos;
                guidPropertySet = propinfoset.guidPropertySet;
            }

            return Tuple.Create(infoCount, infoPtr, guidPropertySet);
        }

        protected override bool ReleaseHandle()
        {
            // NOTE: The SafeHandle class guarantees this will be called exactly once and is non-interrutible.
            IntPtr ptr = base.handle;
            base.handle = IntPtr.Zero;
            if (IntPtr.Zero != ptr)
            {
                int count = this.setCount;
                for (int i = 0; i < count; ++i)
                {
                    int offset = (i * ODB.SizeOf_tagDBPROPINFOSET);
                    IntPtr infoPtr = Marshal.ReadIntPtr(ptr, offset);
                    if (IntPtr.Zero != infoPtr)
                    {
                        int infoCount = Marshal.ReadInt32(ptr, offset + ADP.PtrSize);

                        for (int k = 0; k < infoCount; ++k)
                        {
                            IntPtr valuePtr = ADP.IntPtrOffset(infoPtr, (k * ODB.SizeOf_tagDBPROPINFO) + ODB.OffsetOf_tagDBPROPINFO_Value);
                            SafeNativeMethods.VariantClear(valuePtr);
                        }
                        SafeNativeMethods.CoTaskMemFree(infoPtr); // was allocated by provider
                    }
                }
                SafeNativeMethods.CoTaskMemFree(ptr);
            }

            ptr = this.descBuffer;
            this.descBuffer = IntPtr.Zero;
            if (IntPtr.Zero != ptr)
            {
                SafeNativeMethods.CoTaskMemFree(ptr);
            }
            return true;
        }

        internal static Type FromVtType(int vartype) =>
            (VarEnum)vartype switch
            {
                VarEnum.VT_EMPTY => null,
                VarEnum.VT_NULL => typeof(System.DBNull),
                VarEnum.VT_I2 => typeof(short),
                VarEnum.VT_I4 => typeof(int),
                VarEnum.VT_R4 => typeof(float),
                VarEnum.VT_R8 => typeof(double),
                VarEnum.VT_CY => typeof(decimal),
                VarEnum.VT_DATE => typeof(System.DateTime),
                VarEnum.VT_BSTR => typeof(string),
                VarEnum.VT_DISPATCH => typeof(object),
                VarEnum.VT_ERROR => typeof(int),
                VarEnum.VT_BOOL => typeof(bool),
                VarEnum.VT_VARIANT => typeof(object),
                VarEnum.VT_UNKNOWN => typeof(object),
                VarEnum.VT_DECIMAL => typeof(decimal),
                VarEnum.VT_I1 => typeof(sbyte),
                VarEnum.VT_UI1 => typeof(byte),
                VarEnum.VT_UI2 => typeof(ushort),
                VarEnum.VT_UI4 => typeof(uint),
                VarEnum.VT_I8 => typeof(long),
                VarEnum.VT_UI8 => typeof(ulong),
                VarEnum.VT_INT => typeof(int),
                VarEnum.VT_UINT => typeof(uint),
                _ => typeof(object),
            };
    }
}
