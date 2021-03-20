using System;
using System.Diagnostics;

namespace ArbWeb
{
    // ============================================================================
    // M I C R O  T I M E R
    // ============================================================================
    public class MicroTimer
    {
        private Stopwatch m_sw;

        private int m_msec;
        private int m_msecStop;

        public MicroTimer()
        {
            m_sw = new Stopwatch();
            m_sw.Start();

            m_msec = Environment.TickCount;
            m_msecStop = -1;
        }

        public void Reset()
        {
            m_sw.Reset();
        }

        public void Start()
        {
            m_sw.Start();
        }

        public void Stop()
        {
            m_sw.Stop();
            m_msecStop = Environment.TickCount;
        }

        public double Seconds()
        {
            return m_sw.ElapsedMilliseconds / 1000.0;
            // return (Environment.TickCount - m_msec) / 1000.0;
        }

        public long Microsec
        {
            get
            {
                return m_sw.ElapsedTicks / (Stopwatch.Frequency / (1000L * 1000L));
            }
        }

        public long Msec()
        {
            return m_sw.ElapsedMilliseconds;
            // return m_msecStop - m_msec;
        }

        public double MsecFloat
        {
            get
            {
                return ((double)Microsec) / 1000.0;
            }
        }
        public string Elapsed()
        {
            if (m_sw.IsRunning)
                m_sw.Stop();
            //if (m_msecStop == -1)
            //Stop();

            return $"{Seconds().ToString()}";
        }

        public string ElapsedMsec()
        {
            if (m_sw.IsRunning)
                m_sw.Stop();
            //if (m_msecStop == -1)
            //Stop();

            return Msec().ToString();
        }

    }
}