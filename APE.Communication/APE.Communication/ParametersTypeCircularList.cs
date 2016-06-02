using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace APE.Communication
{
    internal class ParametersTypeCircularList
    {
        Type[][] m_ParametersType;
        Int64[] m_TypeCodeKey;
        int m_OldestItemInList;
        int m_ListSize;

        public void AddToList(Int64 TypeCodeKey, Type[] ParametersType)
        {
            //Add an item to the list replacing the eldest item
            m_TypeCodeKey[m_OldestItemInList] = TypeCodeKey;
            m_ParametersType[m_OldestItemInList] = ParametersType;
            m_OldestItemInList++;
            if (m_OldestItemInList > m_ListSize - 1)
            {
                m_OldestItemInList = 0;
            }
        }

        public void GetFromList(Int64 TypeCodeKey, out Type[] ParametersType)
        {
            //Search for the item starting at the youngest item in the list
            for (int i = m_OldestItemInList - 1; i >= 0; i--)
            {
                if (m_TypeCodeKey[i] == TypeCodeKey)
                {
                    ParametersType = m_ParametersType[i];
                    return;
                }
            }

            for (int i = m_ListSize - 1; i >= m_OldestItemInList; i--)
            {
                if (m_TypeCodeKey[i] == TypeCodeKey)
                {
                    ParametersType = m_ParametersType[i];
                    return;
                }
            }

            ParametersType = null;
        }

        public ParametersTypeCircularList(int Length)
        {
            if (Length < 1)
            {
                throw new Exception("Length must be positive");
            }

            m_ListSize = Length;
            m_ParametersType = new Type[Length][];
            m_TypeCodeKey = new Int64[Length];
        }
    }
}
