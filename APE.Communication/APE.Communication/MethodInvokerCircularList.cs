using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace APE.Communication
{
    internal class MethodInvokerCircularList
    {
        RuntimeTypeHandle[] m_RunTimeTypeHandle;
        string[] m_Name;
        Fasterflect.MethodInvoker[] m_MethodInvoker;
        Int64[] m_TypeCodeKey;
        int m_OldestItemInList;
        int m_ListSize;

        public void AddToList(RuntimeTypeHandle RuntimeTypeHandle, string Name, Int64 TypeCodeKey, Fasterflect.MethodInvoker MethodInvoker)
        {
            //Add an item to the list replacing the eldest item
            m_RunTimeTypeHandle[m_OldestItemInList] = RuntimeTypeHandle;
            m_Name[m_OldestItemInList] = Name;
            m_TypeCodeKey[m_OldestItemInList] = TypeCodeKey;
            m_MethodInvoker[m_OldestItemInList] = MethodInvoker;

            m_OldestItemInList++;
            if (m_OldestItemInList > m_ListSize - 1)
            {
                m_OldestItemInList = 0;
            }
        }

        public void GetFromList(RuntimeTypeHandle RuntimeTypeHandle, string Name, Int64 TypeCodeKey, out Fasterflect.MethodInvoker MethodInvoker)
        {
            //Search for the item starting at the youngest item in the list
            for (int i = m_OldestItemInList - 1; i >= 0; i--)
            {
                if (m_RunTimeTypeHandle[i].Equals(RuntimeTypeHandle) && m_Name[i] == Name && m_TypeCodeKey[i] == TypeCodeKey)
                {
                    MethodInvoker = m_MethodInvoker[i];
                    return;
                }
            }

            for (int i = m_ListSize - 1; i >= m_OldestItemInList; i--)
            {
                if (m_RunTimeTypeHandle[i].Equals(RuntimeTypeHandle) && m_Name[i] == Name && m_TypeCodeKey[i] == TypeCodeKey)
                {
                    MethodInvoker = m_MethodInvoker[i];
                    return;
                }
            }

            MethodInvoker = null;
        }

        public MethodInvokerCircularList(int Length)
        {
            if (Length < 1)
            {
                throw new Exception("Length must be positive");
            }

            m_ListSize = Length;
            m_RunTimeTypeHandle = new RuntimeTypeHandle[Length];
            m_Name = new string[Length];
            m_MethodInvoker = new Fasterflect.MethodInvoker[Length];
            m_TypeCodeKey = new Int64[Length];
        }
    }
}
