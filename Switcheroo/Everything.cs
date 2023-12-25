using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace Switcheroo {
    public class EverythingSearch {
        [DllImport("Everything64.dll", CharSet = CharSet.Unicode)]
        public static extern int Everything_SetSearch(string lpSearchString);

        [DllImport("Everything64.dll")]
        public static extern void Everything_SetMatchPath(bool bEnable);

        [DllImport("Everything64.dll")]
        public static extern void Everything_SetMatchCase(bool bEnable);

        [DllImport("Everything64.dll")]
        public static extern void Everything_SetMatchWholeWord(bool bEnable);

        [DllImport("Everything64.dll")]
        public static extern void Everything_SetRegex(bool bEnable);

        [DllImport("Everything64.dll")]
        public static extern void Everything_SetMax(int dwMax);

        [DllImport("Everything64.dll")]
        public static extern void Everything_SetOffset(int dwOffset);

        [DllImport("Everything64.dll")]
        public static extern bool Everything_Query(bool bWait);

        [DllImport("Everything64.dll", CharSet = CharSet.Unicode)]
        public static extern void Everything_GetResultFullPathName(int nIndex, StringBuilder lpString, int nMaxCount);

        [DllImport("Everything64.dll")]
        public static extern int Everything_GetNumResults();

        public static List<string> SearchFile(string fileWord)
        {
            // 검색 쿼리 설정
            Everything_SetSearch(fileWord);
            Everything_Query(true);
            List<string> result = new List<string>();

            int numResults = Everything_GetNumResults();
            for (int i = 0; i < numResults; i++)
            {
                var sb = new StringBuilder(260);
                Everything_GetResultFullPathName(i, sb, sb.Capacity);
                Trace.WriteLine("Find : [" + sb.ToString() + "]");
                result.Add(sb.ToString());
            }
            return result;
        }
    }
}
