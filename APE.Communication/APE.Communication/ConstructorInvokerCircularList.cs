using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace APE.Communication
{
    internal class ConstructorInvokerCircularList
    {
        string[] m_Name;
        Fasterflect.ConstructorInvoker[] m_ConstructorInvoker;
        Int64[] m_TypeCodeKey;
        int m_OldestItemInList;
        int m_ListSize;

        public void AddToList(string Name, Int64 TypeCodeKey, Fasterflect.ConstructorInvoker ConstructorInvoker)
        {
            //Add an item to the list replacing the eldest item
            m_Name[m_OldestItemInList] = Name;
            m_TypeCodeKey[m_OldestItemInList] = TypeCodeKey;
            m_ConstructorInvoker[m_OldestItemInList] = ConstructorInvoker;

            m_OldestItemInList++;
            if (m_OldestItemInList > m_ListSize - 1)
            {
                m_OldestItemInList = 0;
            }
        }

        public void GetFromList(string Name, Int64 TypeCodeKey, out Fasterflect.ConstructorInvoker ConstructorInvoker)
        {
            //Search for the item starting at the youngest item in the list
            for (int i = m_OldestItemInList - 1; i >= 0; i--)
            {
                if (m_Name[i] == Name && m_TypeCodeKey[i] == TypeCodeKey)
                {
                    ConstructorInvoker = m_ConstructorInvoker[i];
                    return;
                }
            }

            for (int i = m_ListSize - 1; i >= m_OldestItemInList; i--)
            {
                if (m_Name[i] == Name && m_TypeCodeKey[i] == TypeCodeKey)
                {
                    ConstructorInvoker = m_ConstructorInvoker[i];
                    return;
                }
            }

            ConstructorInvoker = null;
        }

        public ConstructorInvokerCircularList(int Length)
        {
            if (Length < 1)
            {
                throw new Exception("Length must be positive");
            }

            m_ListSize = Length;
            m_Name = new string[Length];
            m_ConstructorInvoker = new Fasterflect.ConstructorInvoker[Length];
            m_TypeCodeKey = new Int64[Length];
        }
    }
}
