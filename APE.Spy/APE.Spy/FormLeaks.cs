using APE.Communication;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace APE.Spy
{
    public partial class FormLeaks : Form
    {
        APEIPC m_APE;

        public FormLeaks(APEIPC APE)
        {
            m_APE = APE;
            InitializeComponent();
        }

        List<string> Baseline = new List<string>();
        List<string> Leaks;

        private void buttonBaseline_Click(object sender, EventArgs e)
        {
            if (radioButtonDotNET.Checked)
            {
                m_APE.AddFirstMessageDumpControl();
            }
            else
            {
                m_APE.AddFirstMessageDumpActiveX();
            }
            m_APE.SendMessages(EventSet.APE);
            m_APE.WaitForMessages(EventSet.APE);
            string res = m_APE.GetValueFromMessage();
            if (string.IsNullOrEmpty(res))
            {
                Baseline = new List<string>();
            }
            else
            {
                string[] splitSeparator = { "\r\n" };
                Baseline = res.Split(splitSeparator, StringSplitOptions.None).ToList();
            }
        }

        private void buttonLeaks_Click(object sender, EventArgs e)
        {
            
            if (radioButtonDotNET.Checked)
            {
                m_APE.AddFirstMessageDumpControl();
            }
            else
            {
                m_APE.AddFirstMessageDumpActiveX();
            }
            m_APE.SendMessages(EventSet.APE);
            m_APE.WaitForMessages(EventSet.APE);
            string res = m_APE.GetValueFromMessage();
            if (string.IsNullOrEmpty(res))
            {
                Leaks = new List<string>();
            }
            else
            {
                string[] splitSeparator = { "\r\n" };
                Leaks = res.Split(splitSeparator, StringSplitOptions.None).ToList();
            }

            List<string> leak = new List<string>();
            leak.Add("Leaks:");
            foreach (string potentialLeak in Leaks)
            {
                if (!Baseline.Contains(potentialLeak))
                {
                    leak.Add(potentialLeak);
                }
            }

            textBoxLeaks.Lines = leak.ToArray();
        }

        private void buttonGCFull_Click(object sender, EventArgs e)
        {
            m_APE.AddFirstMessageGarbageCollect(System.GC.MaxGeneration);
            m_APE.SendMessages(EventSet.APE);
            m_APE.WaitForMessages(EventSet.APE);
        }
    }
}
