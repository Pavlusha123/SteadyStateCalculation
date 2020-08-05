using System;
using System.Collections.Generic;
using System.Text;

namespace SteadyStateCalculation
{
    class Line
    {
        public bool LineState { get; set; }
        public int LineStart { get; set; }
        public int LineEnd { get; set; }
        public string LineName { get; set; }
        public Complex LineResistance { get; set; }
        public Complex LineConductivity { get; set; }
        public Complex LineLoadStart { get; set; }
        public Complex LineLoadEnd { get; set; }
        public Complex LineLosses { get; set; }
        public double LineCurrent { get; set; }
        public bool LineCommutation { get; set; }
    }
}
