using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace UntisUPZperMail
{
    class Teachers
    {
        private readonly List<string> teachersListBeforCheck = new List<string>();
        private readonly List<string> teachersListAfterCheck = new List<string>();
        private readonly Dictionary<string, int> keyValuePairs = new Dictionary<string, int>();
        public Teachers(string untisPath, StackPanel stackPanel)
        {
            string teachersLongString = File.ReadAllText(@"G:\Ablage neu\03 Schulverwaltung\SchVwSoftware\config\UntisUPZperMail\Teachers.txt");
            teachersLongString = teachersLongString.Replace("\r\n", "*");
            string[] teacherSingleString = teachersLongString.Split('*');
            foreach (var element in teacherSingleString)
            {
                Debug.Print("BeforeCheck: " + element);
                teachersListBeforCheck.Add(element);
            }
            teachersListBeforCheck.Sort();
            stackPanel.Children.Clear();
            TextBlock textBlock = new TextBlock
            {
                Text = "Folgende Lehrkräfte ohne PDF gefunden:",
                FontSize = 12,
                TextDecorations = TextDecorations.Underline
            };
            stackPanel.Children.Add(textBlock);
            foreach (string element in teachersListBeforCheck)
            {
                string[] teacherElement = element.Split('#');
                if (!File.Exists(string.Format(@"{0}\{1}\{2}.pdf", untisPath, teacherElement[1], teacherElement[0])))
                {
                    Debug.Print("After Check: " + element);
                    textBlock = new TextBlock
                    {
                        Text = teacherElement[0],
                        FontSize = 12
                    };
                    stackPanel.Children.Add(textBlock);
                    teachersListAfterCheck.Add(element);
                }
            }
            if (teachersListAfterCheck.Count == 0)
            {
                stackPanel.Children.Clear();
                textBlock = new TextBlock
                {
                    Text = "Es gibt bereits für jede Lehrkraft ein PDF?",
                    FontSize = 20
                };
                stackPanel.Children.Add(textBlock);
            }
        }

        public void MakePdfDictonary(List<string> pdfSubstring)
        {
            for (int i = 0; i < pdfSubstring.Count(); i++)
            {
                for (int ii = 0; ii < teachersListAfterCheck.Count(); ii++)
                {
                    string[] teacherElement = teachersListAfterCheck[ii].Split('#');
                    if (pdfSubstring[i].Contains(teacherElement[0]))
                    {
                        keyValuePairs.Add(teachersListAfterCheck[ii], i + 1);
                        break;
                    }
                }
            }
        }

        public List<string> ListBeforeCheck => teachersListBeforCheck;
        public List<string> ListAfterCheck => teachersListAfterCheck;
        public Dictionary<string, int> GetDictonary => keyValuePairs;
        public int GetNumberOfListElements(List<string> buffer) => buffer.Count;
        public int GetNumberOfDictonaryElements => keyValuePairs.Count;
    }
}
