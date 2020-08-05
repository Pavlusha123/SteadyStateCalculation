using System;
using System.Collections.Generic;
using System.Text;

namespace SteadyStateCalculation
{
    class LinkWithMatLab
    {
        public static void CallMatLab(string func, string A, string b, string Aeq, string beq, string lb, string ub, string x0, ref Matrix EstimatedVoltage, ref Matrix EstimatedAngle)
        {
            Object[] Comands = {"format long;"+func+A+b+Aeq+beq+lb+ub+x0+"x=fmincon(func,x0,A,b,Aeq,beq,lb,ub)"};
            Type TypeMatLab = Type.GetTypeFromProgID("Matlab.Application");
            Object MatLab = Activator.CreateInstance(TypeMatLab);
            Object result = TypeMatLab.InvokeMember("Execute", System.Reflection.BindingFlags.InvokeMethod, null, MatLab, Comands);
            string str = result.ToString();
            str = str.Replace('.', ',');
            str = str.Substring(str.LastIndexOf('=')+1);
            Char[] sep = { '\n', ' ' };
            var mas = str.Split(sep,StringSplitOptions.RemoveEmptyEntries);
            int j = 0;
            int k = 0;
            bool ColumnIsFound = false;
            for(int i=0;i<mas.Length;i++)
            {
                if (mas[i] == "Columns")
                {
                    ColumnIsFound = true;
                    i += 4;
                }
                if (ColumnIsFound)
                {
                    if (Double.TryParse(mas[i].ToString(), out double value))
                    {
                        if (j < EstimatedVoltage.GetRows) EstimatedVoltage.matrix[j++, 0] = value * 100.0;
                        else
                        {
                            EstimatedAngle.matrix[k++, 0] = value*100.0;
                        }
                    }
                    else Console.WriteLine($"Unable to parse {mas[i]}");
                }
            }
            Console.WriteLine("Расчетные напряжения узлов: \n");
            Console.WriteLine(EstimatedVoltage);
            Console.WriteLine(EstimatedAngle);
        }
    }
}
