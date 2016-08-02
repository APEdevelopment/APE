//
//Copyright 2016 David Beales
//
//Licensed under the Apache License, Version 2.0 (the "License");
//you may not use this file except in compliance with the License.
//You may obtain a copy of the License at
//
//    http://www.apache.org/licenses/LICENSE-2.0
//
//Unless required by applicable law or agreed to in writing, software
//distributed under the License is distributed on an "AS IS" BASIS,
//WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//See the License for the specific language governing permissions and
//limitations under the License.
//
using System;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using NM = APE.Native.NativeMethods;

namespace APE.Communication
{
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    unsafe internal struct Parameters
    {
        private const int MaxParameters = 15;

        public fixed Int32 TypeCode[MaxParameters];
        public fixed Boolean Boolean[MaxParameters];
        public fixed Byte Byte[MaxParameters];
        public fixed Char Char[MaxParameters];
        public fixed Int64 DateTimeBinary[MaxParameters];
        public fixed Int32 DecimalBits0[MaxParameters];
        public fixed Int32 DecimalBits1[MaxParameters];
        public fixed Int32 DecimalBits2[MaxParameters];
        public fixed Int32 DecimalBits3[MaxParameters];
        public fixed Double Double[MaxParameters];
        public fixed Int16 Int16[MaxParameters];
        public fixed Int32 Int32[MaxParameters];
        public fixed Int64 Int64[MaxParameters];
        public fixed SByte SByte[MaxParameters];
        public fixed Single Single[MaxParameters];
        public fixed Int32 StringOffset[MaxParameters];
        public fixed Int32 StringLength[MaxParameters];
        public fixed UInt16 UInt16[MaxParameters];
        public fixed UInt32 UInt32[MaxParameters];
        public fixed UInt64 UInt64[MaxParameters];
        public fixed Int64 IntPtr[MaxParameters];
    }

    public enum ApeTypeCode : int
    {
        Empty = 0,          // Null reference
        Object = 1,         // Instance that isn't a value
        DBNull = 2,         // Database null value
        Boolean = 3,        // Boolean
        Char = 4,           // Unicode character
        SByte = 5,          // Signed 8-bit integer
        Byte = 6,           // Unsigned 8-bit integer
        Int16 = 7,          // Signed 16-bit integer
        UInt16 = 8,         // Unsigned 16-bit integer
        Int32 = 9,          // Signed 32-bit integer
        UInt32 = 10,        // Unsigned 32-bit integer
        Int64 = 11,         // Signed 64-bit integer
        UInt64 = 12,        // Unsigned 64-bit integer
        Single = 13,        // IEEE 32-bit float
        Double = 14,        // IEEE 64-bit double
        Decimal = 15,       // Decimal
        DateTime = 16,      // DateTime
        IntPtr = 17,        // IntPtr
        String = 18,        // Unicode character string
        DataStore = 19,     // APE datastore
        Image = 20,         // Image
    }

    [Flags]
    internal enum MessageAction : int
    {
        RemoveFileMapping = -1,
        None = 0,
        Find = 1,
        Refind = 2,
        ReflectGet = 3,
        GetResult = 4,
        GetListViewGroupRectangle = 5,
        GetListViewItemRectangle = 6,
        SetTimeOuts = 7,
        GetTitleBarItemRectangle = 8,
        AddMouseHook = 9,
        RemoveMouseHook = 10,
        WaitForMouseState = 11,
        GarbageCollect = 12,
        GetContextMenuStrip = 13,
        GetAppDomains = 14,
        GetRecognisedType = 15,
        GetApeTypeFromType = 16,
        GetApeTypeFromObject = 17,
        ReflectPoll = 18,
        ConvertType = 19,
        UnderlyingGridFromResultsGrid = 20,
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    internal struct Message
    {
        public int MoreStringData;
        public int TotalStringDataLength;
        public MessageAction Action;
        public DataStores SourceStore;
        public DataStores DestinationStore;
        public int NameOffset;
        public int NameLength;
        public int NumberOfParameters;
        public MemberTypes MemberType;
        public long TypeCodeKey;
        public Parameters Parameter;
    }

    unsafe public class Parameter
    {
        internal static int OneLargerThanApeTypeCodeEnumMax = (int)Enum.GetValues(typeof(ApeTypeCode)).Cast<ApeTypeCode>().Max() + 1;

        public Parameter(APEIPC instance, Boolean param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.Boolean[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.Boolean;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax) + (int)ApeTypeCode.Boolean);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, Byte param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.Byte[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.Byte;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax) + (int)ApeTypeCode.Byte);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, Char param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.Char[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.Char;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax) + (int)ApeTypeCode.Char);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, DateTime param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.DateTimeBinary[PtrMessage->NumberOfParameters] = param.ToBinary();
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.DateTime;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax) + (int)ApeTypeCode.DateTime);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, Decimal param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));
            int[] DecimalBits = decimal.GetBits(param);

            PtrMessage->Parameter.DecimalBits0[PtrMessage->NumberOfParameters] = DecimalBits[0];
            PtrMessage->Parameter.DecimalBits1[PtrMessage->NumberOfParameters] = DecimalBits[1];
            PtrMessage->Parameter.DecimalBits2[PtrMessage->NumberOfParameters] = DecimalBits[2];
            PtrMessage->Parameter.DecimalBits3[PtrMessage->NumberOfParameters] = DecimalBits[3];

            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.Decimal;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax) + (int)ApeTypeCode.Decimal);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, Double param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.Double[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.Double;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax) + (int)ApeTypeCode.Double);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, Int16 param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.Int16[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.Int16;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax) + (int)ApeTypeCode.Int16);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, Int32 param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.Int32[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.Int32;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax) + (int)ApeTypeCode.Int32);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, DataStores param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.Int32[PtrMessage->NumberOfParameters] = (Int32)param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.DataStore;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax) + (int)ApeTypeCode.DataStore);
            PtrMessage->NumberOfParameters++;
        }


        public Parameter(APEIPC instance, IntPtr param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.IntPtr[PtrMessage->NumberOfParameters] = param.ToInt64();
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.IntPtr;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax) + (int)ApeTypeCode.IntPtr);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, Int64 param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.Int64[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.Int64;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax) + (int)ApeTypeCode.Int64);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, SByte param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.SByte[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.SByte;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax) + (int)ApeTypeCode.SByte);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, Single param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.Single[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.Single;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax) + (int)ApeTypeCode.Single);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, String[] param)
        {
            Parameter p = new Parameter(instance, ConvertString1dArrayToString(param));
            //TODO update the type so we know its a 1d array for the other side to decode
        }

        public Parameter(APEIPC instance, String[,] param)
        {
            Parameter p = new Parameter(instance, ConvertString2dArrayToString(param));
            //TODO update the type so we know its a 2d array for the other side to decode
        }

        private string ConvertString1dArrayToString(string[] StringArray)
        {
            StringBuilder sb = new StringBuilder(StringArray.Length * 4);

            for (int y = 0; y < StringArray.GetUpperBound(0) + 1; y++)
            {
                if (y < StringArray.GetUpperBound(0))
                {
                    sb.Append("\n");
                }
            }
            return sb.ToString();
        }

        private string ConvertString2dArrayToString(string[,] StringArray)
        {
            StringBuilder sb = new StringBuilder(StringArray.Length * 4);

            for (int y = 0; y < StringArray.GetUpperBound(0) + 1; y++)
            {
                for (int x = 0; x < StringArray.GetUpperBound(1) + 1; x++)
                {
                    sb.Append(StringArray[y, x]);
                    if (x < StringArray.GetUpperBound(1))
                    {
                        sb.Append("\t");
                    }
                    else
                    {
                        if (y < StringArray.GetUpperBound(0))
                        {
                            sb.Append("\n");
                        }
                    }
                }
            }
            return sb.ToString();
        }

        public Parameter(APEIPC instance, string param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            if (param != null)
            {
                fixed (void* PtrString = param)
                {
                    NM.CopyMemory(instance.m_IntPtrMemoryMappedFileViewStringStore + instance.m_StringStoreOffset, (IntPtr)PtrString, (UIntPtr)(param.Length * 2));    //UTF16 charcter: For a 4 byte surrogate pair, length actually returns 2 somewhat confusingly although its convenient for us here, so we can just use length * 2
                }
                PtrMessage->Parameter.StringOffset[PtrMessage->NumberOfParameters] = instance.m_StringStoreOffset;
                PtrMessage->Parameter.StringLength[PtrMessage->NumberOfParameters] = param.Length;
                instance.m_StringStoreOffset = instance.m_StringStoreOffset + (param.Length * 2);
            }
            else
            {
                PtrMessage->Parameter.StringOffset[PtrMessage->NumberOfParameters] = -1;
                PtrMessage->Parameter.StringLength[PtrMessage->NumberOfParameters] = -1;
            }

            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.String;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax) + (int)ApeTypeCode.String);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, Image param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            if (param != null)
            {
                //Convert to byte array
                MemoryStream ms = new MemoryStream();
                ImageFormat theImageFormat = param.RawFormat;

                if (theImageFormat.Guid == ImageFormat.MemoryBmp.Guid)  //MemoryBmp don't seem to be able to be saved
                {
                    theImageFormat = ImageFormat.Bmp;
                }

                param.Save(ms, theImageFormat);
                byte[] imageAsByteArray = ms.ToArray();
                
                fixed (void* PtrByteArray = imageAsByteArray)
                {
                    NM.CopyMemory(instance.m_IntPtrMemoryMappedFileViewStringStore + instance.m_StringStoreOffset, (IntPtr)PtrByteArray, (UIntPtr)(imageAsByteArray.Length));
                }
                PtrMessage->Parameter.StringOffset[PtrMessage->NumberOfParameters] = instance.m_StringStoreOffset;
                PtrMessage->Parameter.StringLength[PtrMessage->NumberOfParameters] = imageAsByteArray.Length;
                instance.m_StringStoreOffset = instance.m_StringStoreOffset + imageAsByteArray.Length;
            }
            else
            {
                PtrMessage->Parameter.StringOffset[PtrMessage->NumberOfParameters] = -1;
                PtrMessage->Parameter.StringLength[PtrMessage->NumberOfParameters] = -1;
            }

            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.Image;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax) + (int)ApeTypeCode.Image);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, UInt16 param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.UInt16[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.UInt16;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax) + (int)ApeTypeCode.UInt16);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, UInt32 param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.UInt32[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.UInt32;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax) + (int)ApeTypeCode.UInt32);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, UInt64 param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.UInt64[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.UInt64;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax) + (int)ApeTypeCode.UInt64);
            PtrMessage->NumberOfParameters++;
        }
    }
}
