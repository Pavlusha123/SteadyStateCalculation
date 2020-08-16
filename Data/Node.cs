using System;
using System.Collections.Generic;
using System.Text;

namespace SteadyStateCalculation
{
    class Node
    {
        public bool NodeState { get; set; }
        public string NodeType { get; set; }
        public int NodeNumber { get; set; }
        public string NodeName { get; set; }
        public double NodeNominalVoltage { get; set; }
        public Complex NodeLoad { get; set; }
        public Complex NodeGeneration { get; set; }
        public Complex NodeConductivity { get; set; }
        public Complex NodeCalcVoltage { get; set; }
    }
}
