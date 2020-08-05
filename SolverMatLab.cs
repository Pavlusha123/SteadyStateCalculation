using DocumentFormat.OpenXml.ExtendedProperties;
using System;
using System.Collections.Generic;
using System.Text;

namespace SteadyStateCalculation
{
    class SolverMatLab
    {
        
        public static void GetConductivityMatrix(Node[] nodesArray, Line[] linesArray, int numberNodes, int baseNumber, out Matrix G, out Matrix B, out Matrix Gb, out Matrix Bb, out Matrix ConductivityMatrix, out Matrix baseConductivity)
        {
            //Console.Write("GetConductivityMatrix...");
            int newNumberNodes = numberNodes - 1;

            G = new Matrix(numberNodes, numberNodes);
            B = new Matrix(numberNodes, numberNodes);


            Gb = new Matrix(newNumberNodes);
            Bb = new Matrix(newNumberNodes);


            ConductivityMatrix = new Matrix((numberNodes - 1) * 2, (numberNodes - 1) * 2);
            baseConductivity = new Matrix((numberNodes - 1) * 2);

            #region Ymatrix
            ////Form Y matrix
            //int r = 0;
            //int c = 0;
            //for(int i=0;i<newNumberNodes;i++)
            //{
            //    if ((i == 0) | (i == (newNumberNodes / 2))) r = 0;
            //    if (r == baseNumber) continue;
            //    for(int j=0;j<newNumberNodes;j++)
            //    {
            //        if ((j == 0)|(j==(newNumberNodes/2))) c = 0;
            //        if (c == baseNumber) continue;
            //        if ((i<(newNumberNodes/2)) & (j<(newNumberNodes/2)))
            //        {
            //            if (i == j)
            //            {
            //                ConductivityMatrix.matrix[i, j] = (-1.0) * (FindAllInclusions(nodesArray, linesArray, r)).Real - nodesArray[r].NodeConductivity.Real;
            //            }
            //            else
            //            {
            //                ConductivityMatrix.matrix[i, j] = FindBranches(nodesArray, linesArray, r, c).Real;
            //            }
            //        }
            //        else if ((i < (newNumberNodes / 2)) & (j > (newNumberNodes / 2-1)))
            //        {
            //            if (i == j)
            //            {
            //                ConductivityMatrix.matrix[i, j] = (FindAllInclusions(nodesArray, linesArray, r)).Image - nodesArray[r].NodeConductivity.Image;
            //            }
            //            else
            //            {
            //                ConductivityMatrix.matrix[i, j] = (-1.0)*FindBranches(nodesArray, linesArray, r, c).Image;
            //            }
            //        }
            //        else if ((i > (newNumberNodes / 2-1)) & (j < (newNumberNodes / 2)))
            //        {
            //            if (i == j)
            //            {
            //                ConductivityMatrix.matrix[i, j] = (-1.0) * (FindAllInclusions(nodesArray, linesArray, r)).Real - nodesArray[r].NodeConductivity.Real;
            //            }
            //            else
            //            {
            //                ConductivityMatrix.matrix[i, j] = FindBranches(nodesArray, linesArray, r, c).Real;
            //            }
            //        }
            //        else
            //        {
            //            if (i == j)
            //            {
            //                ConductivityMatrix.matrix[i, j] = (-1.0) * (FindAllInclusions(nodesArray, linesArray, r)).Image - nodesArray[r].NodeConductivity.Image;
            //            }
            //            else
            //            {
            //                ConductivityMatrix.matrix[i, j] = FindBranches(nodesArray, linesArray, r, c).Image;
            //            }
            //        }

            //    }
            //}
            //Console.WriteLine(ConductivityMatrix);
            #endregion

            //Form G matrix
            for (int i = 0; i < numberNodes; i++)
            {
                for (int j = 0; j < numberNodes; j++)
                {
                    if (i == j)
                    {
                        G.matrix[i, j] = (-1) * (FindAllInclusions(nodesArray, linesArray, i)).Real - nodesArray[i].NodeConductivity.Real;
                    }
                    else
                    {
                        G.matrix[i, j] = FindBranches(nodesArray, linesArray, i, j).Real;
                    }
                }
            }

            //Form B matrix
            for (int i = 0; i < numberNodes; i++)
            {
                for (int j = 0; j < numberNodes; j++)
                {
                    if (i == j)
                    {
                        B.matrix[i, j] = (-1) * (FindAllInclusions(nodesArray, linesArray, i)).Image - nodesArray[i].NodeConductivity.Image;
                    }
                    else
                    {
                        B.matrix[i, j] = FindBranches(nodesArray, linesArray, i, j).Image;
                    }
                }
            }

            //Form G base
            int k = 0;
            for (int i = 0; i < numberNodes; i++)
            {
                if (i == baseNumber)
                {
                    continue;
                }
                Gb.matrix[k, 0] = G.matrix[i, baseNumber];
                k++;
            }

            //Form G real
            G = Matrix.CreateMinor(G, baseNumber, baseNumber);

            //Form B base
            k = 0;
            for (int i = 0; i < numberNodes; i++)
            {
                if (i == baseNumber)
                {
                    continue;
                }
                Bb.matrix[k, 0] = B.matrix[i, baseNumber];
                k++;
            }

            //Form B real
            B = Matrix.CreateMinor(B, baseNumber, baseNumber);

            #region Ymatrixworking
            //Form Y matrix
            int r = 0;
            for (int i = 0; i < newNumberNodes; i++)
            {
                int c = 0;
                for (int j = 0; j < newNumberNodes; j++)
                {
                    ConductivityMatrix.matrix[i, j] = G.matrix[r, c];
                    c++;
                }
                r++;
            }
            r = 0;
            for (int i = 0; i < newNumberNodes; i++)
            {
                int c = 0;
                for (int j = newNumberNodes; j < newNumberNodes * 2; j++)
                {
                    ConductivityMatrix.matrix[i, j] = (-1.0) * B.matrix[r, c];
                    c++;
                }
                r++;
            }
            r = 0;
            for (int i = newNumberNodes; i < newNumberNodes * 2; i++)
            {
                int c = 0;
                for (int j = 0; j < newNumberNodes; j++)
                {
                    ConductivityMatrix.matrix[i, j] = B.matrix[r, c];
                    c++;
                }
                r++;
            }
            r = 0;
            for (int i = newNumberNodes; i < newNumberNodes * 2; i++)
            {
                int c = 0;
                for (int j = newNumberNodes; j < newNumberNodes * 2; j++)
                {
                    ConductivityMatrix.matrix[i, j] = G.matrix[r, c];
                    c++;
                }
                r++;
            }

            //Form Yb Matrix
            r = 0;
            for (int i = 0; i < newNumberNodes; i++)
            {
                baseConductivity.matrix[i, 0] = Gb.matrix[r, 0];
                r++;
            }
            r = 0;
            for (int i = newNumberNodes; i < newNumberNodes * 2; i++)
            {
                baseConductivity.matrix[i, 0] = Bb.matrix[r, 0];
                r++;
            }
            //Console.WriteLine(ConductivityMatrix);
            //Console.WriteLine();
            //Console.WriteLine(baseConductivity);
            //Console.WriteLine("Success!");
            #endregion
        }

        public static Complex FindAllInclusions(Node[] nodesArray, Line[] linesArray, int i)
        {
            //Console.Write("FindAllInclusions...");

            Complex Conductivity = new Complex(0, 0);
            Complex One = new Complex(1, 0);

            foreach (var line in linesArray)
            {
                if (line.LineStart == nodesArray[i].NodeNumber | line.LineEnd == nodesArray[i].NodeNumber)
                {
                    if (line.LineState == false)
                    {
                        Conductivity += ((One / (line.LineResistance)) + line.LineConductivity);
                    }
                }
            }
            //Console.WriteLine("Success!");
            return Conductivity;
        }

        public static Complex FindBranches(Node[] nodesArray, Line[] linesArray, int i, int j)
        {
            //Console.Write("FindBranches...");

            Complex Conductivity = new Complex(0, 0);

            foreach (var line in linesArray)
            {
                if (((line.LineStart == nodesArray[i].NodeNumber) & (line.LineEnd == nodesArray[j].NodeNumber)) | ((line.LineStart == nodesArray[j].NodeNumber) & (line.LineEnd == nodesArray[i].NodeNumber)))
                {
                    if (line.LineState == false)
                    {
                        Conductivity += (new Complex(1.0, 0) / (line.LineResistance));
                    }
                }
            }
            //Console.WriteLine("Success!");
            return Conductivity;
        }

        public static void GetLoadMatrix(Node[] nodesArray, int baseNumber, ref Matrix LoadReal, ref Matrix LoadImage, ref Matrix TanFi)
        {
            //Console.Write("GetLoadMatrix...");

            //Form P
            int k = 0;
            for (int i = 0; i < nodesArray.Length; i++)
            {
                if (i == baseNumber)
                {
                    continue;
                }
                LoadReal.matrix[k, 0] = nodesArray[i].NodeLoad.Real - nodesArray[i].NodeGeneration.Real;
                TanFi.matrix[k, 0] = nodesArray[i].NodeLoad.Image / nodesArray[i].NodeLoad.Real;
                k++;
            }

            //Form Q
            k = 0;
            for (int i = 0; i < nodesArray.Length; i++)
            {
                if (i == baseNumber)
                {
                    continue;
                }
                LoadImage.matrix[k, 0] = nodesArray[i].NodeLoad.Image - nodesArray[i].NodeGeneration.Image;
                k++;
            }

            
            //Console.WriteLine("Success!");
        }

        public static void NodesVoltageInitialAprox(Node[] nodesArray, int baseNumber, ref Matrix NodesVoltageModule, ref Matrix NodesVoltageAngle)
        {
            int k = 0;
            for (int i = 0; i < nodesArray.Length; i++)
            {
                if (i == baseNumber) continue;
                NodesVoltageModule.matrix[k, 0] = nodesArray[i].NodeNominalVoltage;
                NodesVoltageAngle.matrix[k, 0] = 0;
                k++;
            }

        }

        public static Matrix NodesVoltage(Matrix NodesVoltageModule, Matrix NodesVoltageAngle)
        {
            Matrix NodesVoltage = new Matrix(NodesVoltageModule.GetRows * 2, 1);

            int j = 0;
            for (int i = 0; i < NodesVoltage.GetRows; i++)
            {
                if (i < NodesVoltage.GetRows / 2)
                    NodesVoltage.matrix[i, 0] = NodesVoltageModule.matrix[i, 0];
                else NodesVoltage.matrix[i, 0] = NodesVoltageAngle.matrix[j++, 0];
            }
            return NodesVoltage;
        }

        public static Matrix LoadMatrix(Matrix LoadReal, Matrix LoadImage)
        {
            Matrix LoadMatrix = new Matrix(LoadReal.GetRows * 2, 1);

            int j = 0;
            for (int i = 0; i < LoadMatrix.GetRows; i++)
            {
                if (i < LoadMatrix.GetRows / 2)
                    LoadMatrix.matrix[i, 0] = LoadReal.matrix[i, 0];
                else
                    LoadMatrix.matrix[i, 0] = LoadImage.matrix[j++, 0];
            }
            return LoadMatrix;
        }

        public static void SolveEquation(Matrix NodesVoltage, Matrix G, Matrix B, Matrix Gb, Matrix Bb, Matrix NodesVoltageModule, Matrix NodesVoltageAngle, Matrix LoadMatrixReal, Matrix LoadMatrixImage, Matrix ConductivityMatrix, Matrix baseConductivity, double baseVoltage, out Matrix EstimatedVoltage, out Matrix EstimatedAngle, out Matrix dP, out Matrix dQ)
        {
            #region VoltageInitialAprox
            StringBuilder x0 = new StringBuilder("x0 = [");
            for(int i=0;i<NodesVoltage.GetRows;i++)
            {
                x0.Append($"{NodesVoltage.matrix[i, 0]},");
            }
            x0.Remove(x0.Length-1,1);
            x0.Append("];");
            #endregion
            
            #region InequalityConstraints
            StringBuilder A = new StringBuilder("A = [];");
            StringBuilder b = new StringBuilder("b = [];");
            #endregion

            #region EqualityConstraints
            StringBuilder Aeq = new StringBuilder("Aeq = [];");
            StringBuilder beq = new StringBuilder("beq = [];");
            #endregion

            #region LowerBound
            StringBuilder lb = new StringBuilder("lb = [");
            for (int i = 0; i < NodesVoltageModule.GetRows; i++)
            {
                lb.Append($"450,");
            }
            for (int i = 0; i < NodesVoltageAngle.GetRows; i++)
            {
                lb.Append("-90,");
            }
            lb.Remove(lb.Length - 1, 1);
            lb.Append("];");
            #endregion

            #region UpperBound
            StringBuilder ub = new StringBuilder("ub = [");
            for (int i = 0; i < NodesVoltageModule.GetRows; i++)
            {
                ub.Append($"550,");
            }
            for (int i = 0; i < NodesVoltageAngle.GetRows; i++)
            {
                ub.Append("90,");
            }
            ub.Remove(ub.Length - 1, 1);
            ub.Append("];");
            #endregion

            #region FunctionForOptimization
            StringBuilder func = new StringBuilder("func = @(x)");
            int k;
            int kj;
            for(int i=0;i<NodesVoltageModule.GetRows;i++) //Квадраты невязок активной мощности
            {
                k = i + NodesVoltageModule.GetRows+1;
                func.Append($"(x({i + 1})*(");
                for(int j=0;j<NodesVoltageModule.GetRows;j++) //k kj i+1 j+1
                {
                    kj = j + NodesVoltageModule.GetRows+1;
                    func.Append($"(({G.matrix[i,j]})*cos(x({kj})-x({k}))-({B.matrix[i,j]})*sin(x({kj})-x({k})))*x({j+1})+");
                }
                func.Remove(func.Length - 1, 1);
                func.Append(")");
                func.Append($"+x({i+1})*({baseVoltage})*(({Gb.matrix[i,0]})*cos(x({k}))+({Bb.matrix[i,0]})*sin(x({k})))-({LoadMatrixReal.matrix[i,0]}))^2+");
            }

            for (int i = 0; i < NodesVoltageModule.GetRows; i++) //Квадраты невязок реактивной мощности
            {
                k = i + NodesVoltageModule.GetRows+1;
                func.Append($"(x({i + 1})*(");
                for (int j = 0; j < NodesVoltageModule.GetRows; j++) //k kj i+1 j+1
                {
                    kj = j + NodesVoltageModule.GetRows+1;
                    func.Append($"(({G.matrix[i, j]})*sin(x({kj})-x({k}))+({B.matrix[i, j]})*cos(x({kj})-x({k})))*x({j+1})+");
                }
                func.Remove(func.Length - 1, 1);
                func.Append(")");
                func.Append($"+x({i+1})*({baseVoltage})*((-{Gb.matrix[i, 0]})*sin(x({k}))+({Bb.matrix[i, 0]})*cos(x({k})))+({LoadMatrixImage.matrix[i, 0]}))^2+");
            }
            func.Remove(func.Length - 1, 1);
            func.Append(";");
            func.Replace(',', '.');
            #endregion
            //Console.WriteLine(func);

            #region Calculations
            EstimatedVoltage = new Matrix(NodesVoltageModule.GetRows);
            EstimatedAngle = new Matrix(NodesVoltageAngle.GetRows);
            LinkWithMatLab.CallMatLab(func.ToString(), A.ToString(), b.ToString(), Aeq.ToString(), beq.ToString(), lb.ToString(), ub.ToString(), x0.ToString(), ref EstimatedVoltage, ref EstimatedAngle);

            dP = new Matrix(LoadMatrixReal.GetRows);
            dQ = new Matrix(LoadMatrixImage.GetRows);
            for (int i = 0; i < EstimatedVoltage.GetRows; i++) //Невязки активной мощности
            {
                double temp = 0;
                dP.matrix[i, 0] = EstimatedVoltage.matrix[i, 0];
                for (int j = 0; j < EstimatedVoltage.GetRows; j++) //k kj i+1 j+1
                {
                    temp += (G.matrix[i,j]*Math.Cos(EstimatedAngle.matrix[j,0]-EstimatedAngle.matrix[i,0])-B.matrix[i,j]*
                        Math.Sin(EstimatedAngle.matrix[j,0]-EstimatedAngle.matrix[i,0]))*EstimatedVoltage.matrix[j,0];
                }
                dP.matrix[i, 0] *= temp;
                dP.matrix[i, 0] += EstimatedVoltage.matrix[i, 0] * baseVoltage * (Gb.matrix[i,0]*Math.Cos(EstimatedAngle.matrix[i,0])+Bb.matrix[i,0]*
                    Math.Sin(EstimatedAngle.matrix[i,0]))-LoadMatrixReal.matrix[i,0];
            }

            for (int i = 0; i < EstimatedVoltage.GetRows; i++) //Невязки реактивной мощности
            {
                double temp = 0;
                dQ.matrix[i, 0] = EstimatedVoltage.matrix[i, 0];
                for (int j = 0; j < EstimatedVoltage.GetRows; j++) //k kj i+1 j+1
                {
                    temp += (G.matrix[i, j] * Math.Sin(EstimatedAngle.matrix[j, 0] - EstimatedAngle.matrix[i, 0]) + B.matrix[i, j] * 
                        Math.Cos(EstimatedAngle.matrix[j, 0] - EstimatedAngle.matrix[i, 0])) * EstimatedVoltage.matrix[j, 0];
                }
                dQ.matrix[i, 0] *= temp;
                dQ.matrix[i, 0] += EstimatedVoltage.matrix[i, 0] * baseVoltage * ((-1.0)*Gb.matrix[i, 0] * Math.Sin(EstimatedAngle.matrix[i, 0]) + Bb.matrix[i, 0] * 
                    Math.Cos(EstimatedAngle.matrix[i, 0])) + LoadMatrixImage.matrix[i, 0];
            }
            //Console.WriteLine("\nНевязки:\nАктивной мощности:\n"+dP+"Реактивной мощности:\n"+dQ);
            #endregion
        }

    }

}

