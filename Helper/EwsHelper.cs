using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XEws.Helper
{
    public static class EwsHelper
    {
        /// <summary>
        /// Returns string representation of byte array.
        /// </summary>
        /// <param name="byteArray">byte array to convert to string.</param>
        /// <returns></returns>
        public static string GetStringFromByte(byte[] byteArray)
        {
            return Encoding.UTF8.GetString(byteArray);
        }
    }
}
