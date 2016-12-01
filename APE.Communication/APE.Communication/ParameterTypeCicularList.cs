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
    internal class ParameterTypeCircularList
    {
        Type[] m_ParameterType;
        string[] m_ParameterTypeText;
        int m_OldestItemInList;
        int m_ListSize;

        public void AddToList(string parameterTypeText, Type parametersType)
        {
            //Add an item to the list replacing the eldest item
            m_ParameterTypeText[m_OldestItemInList] = parameterTypeText;
            m_ParameterType[m_OldestItemInList] = parametersType;
            m_OldestItemInList++;
            if (m_OldestItemInList > m_ListSize - 1)
            {
                m_OldestItemInList = 0;
            }
        }

        public void GetFromList(string parameterTypeText, out Type ParametersType)
        {
            //Search for the item starting at the youngest item in the list
            for (int i = m_OldestItemInList - 1; i >= 0; i--)
            {
                if (m_ParameterTypeText[i] == parameterTypeText)
                {
                    ParametersType = m_ParameterType[i];
                    return;
                }
            }

            for (int i = m_ListSize - 1; i >= m_OldestItemInList; i--)
            {
                if (m_ParameterTypeText[i] == parameterTypeText)
                {
                    ParametersType = m_ParameterType[i];
                    return;
                }
            }

            ParametersType = null;
        }

        public ParameterTypeCircularList(int Length)
        {
            if (Length < 1)
            {
                throw new Exception("Length must be positive");
            }

            m_ListSize = Length;
            m_ParameterType = new Type[Length];
            m_ParameterTypeText = new string[Length];
        }
    }
}
