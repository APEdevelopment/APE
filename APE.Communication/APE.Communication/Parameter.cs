using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Reflection;
using NM = APE.Native.NativeMethods;

namespace APE.Communication
{
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    unsafe internal struct Parameters
    {
        private const int MaxParameters = 10;

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
        internal const int OneLargerThanTypeCodeEnumMax = 20;

        public Parameter(APEIPC instance, Boolean param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.Boolean[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.Boolean;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.Boolean);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, Byte param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.Byte[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.Byte;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.Byte);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, Char param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.Char[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.Char;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.Char);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, DateTime param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.DateTimeBinary[PtrMessage->NumberOfParameters] = param.ToBinary();
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.DateTime;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.DateTime);
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

            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.Decimal;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.Decimal);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, Double param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.Double[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.Double;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.Double);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, Int16 param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.Int16[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.Int16;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.Int16);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, Int32 param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.Int32[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.Int32;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.Int32);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, DataStores param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.Int32[PtrMessage->NumberOfParameters] = (Int32)param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = 19;    //19 is unused so we steal it
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + 19);
            PtrMessage->NumberOfParameters++;
        }


        public Parameter(APEIPC instance, IntPtr param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.IntPtr[PtrMessage->NumberOfParameters] = param.ToInt64();
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = 17;    //17 is unused so we steal it
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + 17);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, Int64 param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.Int64[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.Int64;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.Int64);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, SByte param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.SByte[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.SByte;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.SByte);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, Single param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.Single[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.Single;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.Single);
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

        public Parameter(APEIPC instance, String param)
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

            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.String;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.String);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, UInt16 param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.UInt16[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.UInt16;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.UInt16);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, UInt32 param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.UInt32[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.UInt32;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.UInt32);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, UInt64 param)
        {
            Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

            PtrMessage->Parameter.UInt64[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.UInt64;
            PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.UInt64);
            PtrMessage->NumberOfParameters++;
        }
    }
}
