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
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using NM = APE.Native.NativeMethods;

namespace APE.Communication
{
    public enum ParameterType : int
    {
        In = 1,
        Out,
        Ref,
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    unsafe internal struct Parameters
    {
        private const int MaxParameters = 20;

        public fixed Int32 TypeCode[MaxParameters];
        public fixed int ParameterType[MaxParameters];
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

    unsafe public class Parameter
    {
        internal static int OneLargerThanApeTypeCodeEnumMax = (int)Enum.GetValues(typeof(ApeTypeCode)).Cast<ApeTypeCode>().Max() + 1;

        public Parameter(APEIPC instance, Boolean param)
        {
            Message* PtrMessage = instance.GetPointerToNextMessage();
            PtrMessage->Parameter.Boolean[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.Boolean;
            PtrMessage->Parameter.ParameterType[PtrMessage->NumberOfParameters] = (int)ParameterType.In;
            PtrMessage->TypeCodeKey += ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + (int)ApeTypeCode.Boolean) + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + OneLargerThanApeTypeCodeEnumMax + (int)ParameterType.In);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, Byte param)
        {
            Message* PtrMessage = instance.GetPointerToNextMessage();
            PtrMessage->Parameter.Byte[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.Byte;
            PtrMessage->Parameter.ParameterType[PtrMessage->NumberOfParameters] = (int)ParameterType.In;
            PtrMessage->TypeCodeKey += ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + (int)ApeTypeCode.Byte) + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + OneLargerThanApeTypeCodeEnumMax + (int)ParameterType.In);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, Char param)
        {
            Message* PtrMessage = instance.GetPointerToNextMessage();
            PtrMessage->Parameter.Char[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.Char;
            PtrMessage->Parameter.ParameterType[PtrMessage->NumberOfParameters] = (int)ParameterType.In;
            PtrMessage->TypeCodeKey += ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + (int)ApeTypeCode.Char) + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + OneLargerThanApeTypeCodeEnumMax + (int)ParameterType.In);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, DateTime param)
        {
            Message* PtrMessage = instance.GetPointerToNextMessage();
            PtrMessage->Parameter.DateTimeBinary[PtrMessage->NumberOfParameters] = param.ToBinary();
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.DateTime;
            PtrMessage->Parameter.ParameterType[PtrMessage->NumberOfParameters] = (int)ParameterType.In;
            PtrMessage->TypeCodeKey += ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + (int)ApeTypeCode.DateTime) + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + OneLargerThanApeTypeCodeEnumMax + (int)ParameterType.In);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, Decimal param)
        {
            Message* PtrMessage = instance.GetPointerToNextMessage();
            int[] DecimalBits = decimal.GetBits(param);

            PtrMessage->Parameter.DecimalBits0[PtrMessage->NumberOfParameters] = DecimalBits[0];
            PtrMessage->Parameter.DecimalBits1[PtrMessage->NumberOfParameters] = DecimalBits[1];
            PtrMessage->Parameter.DecimalBits2[PtrMessage->NumberOfParameters] = DecimalBits[2];
            PtrMessage->Parameter.DecimalBits3[PtrMessage->NumberOfParameters] = DecimalBits[3];

            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.Decimal;
            PtrMessage->Parameter.ParameterType[PtrMessage->NumberOfParameters] = (int)ParameterType.In;
            PtrMessage->TypeCodeKey += ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + (int)ApeTypeCode.Decimal) + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + OneLargerThanApeTypeCodeEnumMax + (int)ParameterType.In);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, Double param)
        {
            Message* PtrMessage = instance.GetPointerToNextMessage();
            PtrMessage->Parameter.Double[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.Double;
            PtrMessage->Parameter.ParameterType[PtrMessage->NumberOfParameters] = (int)ParameterType.In;
            PtrMessage->TypeCodeKey += ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + (int)ApeTypeCode.Double) + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + OneLargerThanApeTypeCodeEnumMax + (int)ParameterType.In);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, Int16 param)
        {
            Message* PtrMessage = instance.GetPointerToNextMessage();
            PtrMessage->Parameter.Int16[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.Int16;
            PtrMessage->Parameter.ParameterType[PtrMessage->NumberOfParameters] = (int)ParameterType.In;
            PtrMessage->TypeCodeKey += ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + (int)ApeTypeCode.Int16) + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + OneLargerThanApeTypeCodeEnumMax + (int)ParameterType.In);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, Int32 param)
        {
            Message* PtrMessage = instance.GetPointerToNextMessage();
            PtrMessage->Parameter.Int32[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.Int32;
            PtrMessage->Parameter.ParameterType[PtrMessage->NumberOfParameters] = (int)ParameterType.In;
            PtrMessage->TypeCodeKey += ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + (int)ApeTypeCode.Int32) + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + OneLargerThanApeTypeCodeEnumMax + (int)ParameterType.In);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, DataStores param)
        {
            Message* PtrMessage = instance.GetPointerToNextMessage();
            PtrMessage->Parameter.Int32[PtrMessage->NumberOfParameters] = (Int32)param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.DataStore;
            PtrMessage->Parameter.ParameterType[PtrMessage->NumberOfParameters] = (int)ParameterType.In;
            PtrMessage->TypeCodeKey += ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + (int)ApeTypeCode.DataStore) + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + OneLargerThanApeTypeCodeEnumMax + (int)ParameterType.In);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, DataStores param, ParameterType apeParameterType)
        {
            if (apeParameterType == ParameterType.Out)
            {
                throw new Exception("Must specifiy the parameter type when using out datastore parameters");
            }

            Message* PtrMessage = instance.GetPointerToNextMessage();
            PtrMessage->Parameter.Int32[PtrMessage->NumberOfParameters] = (Int32)param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.DataStore;
            PtrMessage->Parameter.ParameterType[PtrMessage->NumberOfParameters] = (int)apeParameterType;
            PtrMessage->TypeCodeKey += ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + (int)ApeTypeCode.DataStore) + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + OneLargerThanApeTypeCodeEnumMax + (int)apeParameterType);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, DataStores param, ParameterType apeParameterType, string parameterType)
        {
            Message* PtrMessage = instance.GetPointerToNextMessage();
            PtrMessage->Parameter.Int32[PtrMessage->NumberOfParameters] = (Int32)param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.DataStore;
            PtrMessage->Parameter.ParameterType[PtrMessage->NumberOfParameters] = (int)apeParameterType;
            PtrMessage->TypeCodeKey += ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + (int)ApeTypeCode.DataStore) + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + OneLargerThanApeTypeCodeEnumMax + (int)apeParameterType);

            if (apeParameterType == ParameterType.Out)
            {
                int length = parameterType.Length * 2;    //UTF16 charcter: For a 4 byte surrogate pair, length actually returns 2 somewhat confusingly although its convenient for us here, so we can just use length * 2
                CheckSize(instance, length);
                fixed (void* PtrString = parameterType)
                {
                    NM.CopyMemory(instance.m_IntPtrMemoryMappedFileViewStringStore + instance.m_StringStoreOffset, (IntPtr)PtrString, (UIntPtr)(length));
                }
                PtrMessage->Parameter.StringOffset[PtrMessage->NumberOfParameters] = instance.m_StringStoreOffset;
                PtrMessage->Parameter.StringLength[PtrMessage->NumberOfParameters] = parameterType.Length;
                instance.m_StringStoreOffset = instance.m_StringStoreOffset + (parameterType.Length * 2);
            }

            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, IntPtr param)
        {
            Message* PtrMessage = instance.GetPointerToNextMessage();
            PtrMessage->Parameter.IntPtr[PtrMessage->NumberOfParameters] = param.ToInt64();
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.IntPtr;
            PtrMessage->Parameter.ParameterType[PtrMessage->NumberOfParameters] = (int)ParameterType.In;
            PtrMessage->TypeCodeKey += ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + (int)ApeTypeCode.IntPtr) + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + OneLargerThanApeTypeCodeEnumMax + (int)ParameterType.In);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, Int64 param)
        {
            Message* PtrMessage = instance.GetPointerToNextMessage();
            PtrMessage->Parameter.Int64[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.Int64;
            PtrMessage->Parameter.ParameterType[PtrMessage->NumberOfParameters] = (int)ParameterType.In;
            PtrMessage->TypeCodeKey += ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + (int)ApeTypeCode.Int64) + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + OneLargerThanApeTypeCodeEnumMax + (int)ParameterType.In);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, SByte param)
        {
            Message* PtrMessage = instance.GetPointerToNextMessage();
            PtrMessage->Parameter.SByte[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.SByte;
            PtrMessage->Parameter.ParameterType[PtrMessage->NumberOfParameters] = (int)ParameterType.In;
            PtrMessage->TypeCodeKey += ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + (int)ApeTypeCode.SByte) + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + OneLargerThanApeTypeCodeEnumMax + (int)ParameterType.In);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, Single param)
        {
            Message* PtrMessage = instance.GetPointerToNextMessage();
            PtrMessage->Parameter.Single[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.Single;
            PtrMessage->Parameter.ParameterType[PtrMessage->NumberOfParameters] = (int)ParameterType.In;
            PtrMessage->TypeCodeKey += ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + (int)ApeTypeCode.Single) + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + OneLargerThanApeTypeCodeEnumMax + (int)ParameterType.In);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, UInt16 param)
        {
            Message* PtrMessage = instance.GetPointerToNextMessage();
            PtrMessage->Parameter.UInt16[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.UInt16;
            PtrMessage->Parameter.ParameterType[PtrMessage->NumberOfParameters] = (int)ParameterType.In;
            PtrMessage->TypeCodeKey += ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + (int)ApeTypeCode.UInt16) + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + OneLargerThanApeTypeCodeEnumMax + (int)ParameterType.In);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, UInt16 param, ParameterType apeParameterType)
        {
            Message* PtrMessage = instance.GetPointerToNextMessage();
            PtrMessage->Parameter.UInt16[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.UInt16;
            PtrMessage->Parameter.ParameterType[PtrMessage->NumberOfParameters] = (int)apeParameterType;
            PtrMessage->TypeCodeKey += ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + (int)ApeTypeCode.UInt16) + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + OneLargerThanApeTypeCodeEnumMax + (int)apeParameterType);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, UInt32 param)
        {
            Message* PtrMessage = instance.GetPointerToNextMessage();
            PtrMessage->Parameter.UInt32[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.UInt32;
            PtrMessage->Parameter.ParameterType[PtrMessage->NumberOfParameters] = (int)ParameterType.In;
            PtrMessage->TypeCodeKey += ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + (int)ApeTypeCode.UInt32) + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + OneLargerThanApeTypeCodeEnumMax + (int)ParameterType.In);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, UInt32 param, ParameterType apeParameterType)
        {
            Message* PtrMessage = instance.GetPointerToNextMessage();
            PtrMessage->Parameter.UInt32[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.UInt32;
            PtrMessage->Parameter.ParameterType[PtrMessage->NumberOfParameters] = (int)apeParameterType;
            PtrMessage->TypeCodeKey += ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + (int)ApeTypeCode.UInt32) + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + OneLargerThanApeTypeCodeEnumMax + (int)apeParameterType);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, UInt64 param)
        {
            Message* PtrMessage = instance.GetPointerToNextMessage();
            PtrMessage->Parameter.UInt64[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.UInt64;
            PtrMessage->Parameter.ParameterType[PtrMessage->NumberOfParameters] = (int)ParameterType.In;
            PtrMessage->TypeCodeKey += ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + (int)ApeTypeCode.UInt64) + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + OneLargerThanApeTypeCodeEnumMax + (int)ParameterType.In);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, UInt64 param, ParameterType apeParameterType)
        {
            Message* PtrMessage = instance.GetPointerToNextMessage();
            PtrMessage->Parameter.UInt64[PtrMessage->NumberOfParameters] = param;
            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.UInt64;
            PtrMessage->Parameter.ParameterType[PtrMessage->NumberOfParameters] = (int)apeParameterType;
            PtrMessage->TypeCodeKey += ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + (int)ApeTypeCode.UInt64) + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + OneLargerThanApeTypeCodeEnumMax + (int)apeParameterType);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, stdole.Picture param)
        {
            Image pictureAsImage = instance.OlePictureToImage(param);
            Parameter paramImage = new Parameter(instance, pictureAsImage);
        }

        public Parameter(APEIPC instance, Image param)
        {
            Message* PtrMessage = instance.GetPointerToNextMessage();
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

                int length = imageAsByteArray.Length;
                CheckSize(instance, length);
                fixed (void* PtrByteArray = imageAsByteArray)
                {
                    NM.CopyMemory(instance.m_IntPtrMemoryMappedFileViewStringStore + instance.m_StringStoreOffset, (IntPtr)PtrByteArray, (UIntPtr)(length));
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
            PtrMessage->Parameter.ParameterType[PtrMessage->NumberOfParameters] = (int)ParameterType.In;
            PtrMessage->TypeCodeKey += ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + (int)ApeTypeCode.Image) + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + OneLargerThanApeTypeCodeEnumMax + (int)ParameterType.In);
            PtrMessage->NumberOfParameters++;
        }

        public Parameter(APEIPC instance, string param)
        {
            Message* PtrMessage = instance.GetPointerToNextMessage();
            if (param != null)
            {
                int length = param.Length * 2;  //UTF16 charcter: For a 4 byte surrogate pair, length actually returns 2 somewhat confusingly although its convenient for us here, so we can just use length * 2
                CheckSize(instance, length);
                fixed (void* PtrString = param)
                {
                    NM.CopyMemory(instance.m_IntPtrMemoryMappedFileViewStringStore + instance.m_StringStoreOffset, (IntPtr)PtrString, (UIntPtr)(length));
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
            PtrMessage->Parameter.ParameterType[PtrMessage->NumberOfParameters] = (int)ParameterType.In;
            PtrMessage->TypeCodeKey += ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + (int)ApeTypeCode.String) + ((PtrMessage->NumberOfParameters * OneLargerThanApeTypeCodeEnumMax * 2) + OneLargerThanApeTypeCodeEnumMax + (int)ParameterType.In);
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
                            sb.Append("\r");
                        }
                    }
                }
            }
            return sb.ToString();
        }

        private void CheckSize(APEIPC instance, int length)
        {
            if (instance.m_StringStoreOffset + length > APEIPC.StringSpaceBytes)
            {
                throw new Exception("Message larger than space available");
            }
        }
    }
}
