using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Fusk
{
    class Program
    {
        static void Main(string[] args)
        {
            List<string> hej = new List<string> { "Jeg", "hedder", "Kaj" };

            string daw = hej.Aggregate(new StringBuilder(),
                  (sb, a) => sb.AppendLine(String.Join(",", a)),
                  sb => sb.ToString());
        }
    }
}
