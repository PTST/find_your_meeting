using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Drawing;


namespace FindMeetingConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            Dictionary<string, List<int>> locations;
            using (System.IO.StreamReader r = new System.IO.StreamReader(@"M:\GitHub\find_your_meeting\resources\locations.json"))
            {
                locations = JsonConvert.DeserializeObject<Dictionary<string, List<int>>>(r.ReadToEnd());
            }

            foreach (string key in locations.Keys)
            {
                Console.WriteLine(key);
            }

            int cordX = locations["TESTLOCATION"][0];
            int cordY = locations["TESTLOCATION"][1];

        Bitmap img = new Bitmap(@"M:\basement.png");


            for (int i = 0; i < 15; i++)
            {
                for (int n = 0; n < 15; n++)
                {
                    img.SetPixel(cordX+i, cordY+n, Color.OrangeRed);
                }
            }

            img.Save(@"m:\test2.png");

            System.Diagnostics.Process.Start(@"m:\test2.png");
        }
    }
}
