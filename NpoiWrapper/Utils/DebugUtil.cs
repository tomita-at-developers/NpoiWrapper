using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace Developers.NpoiWrapper.Utils
{
    internal class DebugTimer
    {
        public Stopwatch Stopwatch= new Stopwatch();
        public void Start()
        {
            Stopwatch.Start();
        }
        public void Stop()
        {
            Stopwatch.Stop();
        }
        public TimeSpan Elapsed
        {
            get { return Stopwatch.Elapsed; }
        }
        public double ElapsedSeconds
        {
            get
            {
                return (double)Stopwatch.ElapsedMilliseconds / 1000;
            }
        }
        public string ElapsedSecondsString
        {
            get
            {
                return ElapsedSeconds.ToString("0.000");
            }
        }
        public string ElapsedString
        {
            get
            {
                return ElapsedSeconds.ToString(@"mm:ss\.fff");
            }
        }
    }
}
