using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SSMSTextResultExcelLib
{
    public static class StringExtension
    {
        public static string RemovePostNewLine(this string inputString)
        {
            string tempString = inputString;
            while (1 == 1)
            {
                if (tempString.EndsWith(Environment.NewLine))
                {
                    //tempString = tempString.Substring(0, tempString.Length - 2);
                    tempString = tempString.Remove(tempString.Length - Environment.NewLine.Length);
                } else
                {
                    break;
                }
            }
            return tempString;
        }
    }
}
