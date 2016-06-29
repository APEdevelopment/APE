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

namespace APE.Communication
{
    internal class MemberGetterCircularList
    {
        IntPtr[] m_SourceTypeHandle;
        string[] m_Name;
        Fasterflect.MemberGetter[] m_MemberGetter;
        int m_OldestItemInList;
        int m_ListSize;

        public void AddToList(IntPtr sourceTypeHandle, string name, Fasterflect.MemberGetter memberGetter)
        {
            //Add an item to the list replacing the eldest item
            m_SourceTypeHandle[m_OldestItemInList] = sourceTypeHandle;
            m_Name[m_OldestItemInList] = name;
            m_MemberGetter[m_OldestItemInList] = memberGetter;

            m_OldestItemInList++;
            if (m_OldestItemInList > m_ListSize - 1)
            {
                m_OldestItemInList = 0;
            }
        }

        public void GetFromList(IntPtr sourceTypeHandle, string name, out Fasterflect.MemberGetter memberGetter)
        {
            //Search for the item starting at the youngest item in the list
            for (int i = m_OldestItemInList - 1; i >= 0; i--)
            {
                if (m_SourceTypeHandle[i] == sourceTypeHandle && m_Name[i] == name)
                {
                    memberGetter = m_MemberGetter[i];
                    return;
                }
            }

            for (int i = m_ListSize - 1; i >= m_OldestItemInList; i--)
            {
                if (m_SourceTypeHandle[i] == sourceTypeHandle && m_Name[i] == name)
                {
                    memberGetter = m_MemberGetter[i];
                    return;
                }
            }

            memberGetter = null;
        }

        public MemberGetterCircularList(int length)
        {
            if (length < 1)
            {
                throw new Exception("Length must be positive");
            }

            m_ListSize = length;
            m_SourceTypeHandle = new IntPtr[length];
            m_Name = new string[length];
            m_MemberGetter = new Fasterflect.MemberGetter[length];
        }
    }
}
