using DocumentFormat.OpenXml.Math;
using System;
using System.Collections.Generic;
using System.Text;

namespace SteadyStateCalculation
{
    abstract class SolverComplex
    {
        public static MatrixComplex GetConductivityMatrix(Node[] nodesArray, Line[] linesArray, int numberNodes, ref MatrixComplex baseConductivity, ref int baseNumber)
        {
            //Console.Write("GetConductivityMatrix...");
            MatrixComplex ConductivityMatrix = new MatrixComplex(numberNodes, numberNodes);

            for (int i = 0; i < numberNodes; i++)
            {
                if (nodesArray[i].NodeType == "База")
                {
                    baseNumber = i;
                }
                for (int j = 0; j < numberNodes; j++)
                {
                    ConductivityMatrix.matrix[i, j] = new Complex();
                    if (i == j)
                    {
                        ConductivityMatrix.matrix[i, j] = (-1) * (FindAllInclusions(nodesArray, linesArray, i)) - nodesArray[i].NodeConductivity;
                    }
                    else
                    {
                        ConductivityMatrix.matrix[i, j] = FindBranches(nodesArray, linesArray, i, j);
                    }
                }
            }
            Console.WriteLine("Полная матрица проводимостей:\n"+ConductivityMatrix);
            int k = 0;
            for(int i =0;i<numberNodes;i++)
            {
                if(i==baseNumber)
                {
                    continue;
                }
                baseConductivity.matrix[k, 0] = ConductivityMatrix.matrix[i,baseNumber];
                k++;
            }
            //Console.WriteLine("Success!");
            return MatrixComplex.CreateMinor(ConductivityMatrix,baseNumber,baseNumber);
        }

        public static Complex FindAllInclusions(Node[] nodesArray, Line[] linesArray, int i)
        {
            //Console.Write("FindAllInclusions...");

            Complex Conductivity = new Complex(0,0);
            Complex One = new Complex(1, 0);
            
            foreach(var line in linesArray)
            {
                if (line.LineStart==nodesArray[i].NodeNumber | line.LineEnd== nodesArray[i].NodeNumber)
                {
                    if(line.LineState==false)
                    {
                        Conductivity += ((One / line.LineResistance) + line.LineConductivity);
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
            Complex One = new Complex(1, 0);

            foreach (var line in linesArray)
            {
                if (((line.LineStart == nodesArray[i].NodeNumber) & (line.LineEnd == nodesArray[j].NodeNumber)) | ((line.LineStart == nodesArray[j].NodeNumber) & (line.LineEnd == nodesArray[i].NodeNumber)))
                {
                    if (line.LineState == false)
                    {
                        Conductivity += (One / line.LineResistance);
                    }
                }
            }
            //Console.WriteLine("Success!");
            return Conductivity;
        }

        public static MatrixComplex GetLoadMatrix(Node[] nodesArray,int baseNumber)
        {
            //Console.Write("GetLoadMatrix...");
            MatrixComplex LoadMatrix = new MatrixComplex(nodesArray.Length-1,1);
            int k = 0;
            for (int i = 0; i < nodesArray.Length; i++)
            {
                if (i == baseNumber)
                {
                    continue;
                }
                LoadMatrix.matrix[k, 0] = nodesArray[i].NodeLoad - nodesArray[i].NodeGeneration;
                k++;
            }
            //Console.WriteLine("Success!");
            return LoadMatrix;
        }

        public static MatrixComplex CalcJacobiMatrix(MatrixComplex ConductivityMatrix, MatrixComplex baseConductivity, MatrixComplex NodesVoltage,MatrixComplex diagConjugateNodesVoltage, Complex baseVoltage)
        {
            MatrixComplex JacobiMatrix = new MatrixComplex(ConductivityMatrix.GetRows,ConductivityMatrix.GetColumns);

            for(int i=0;i<JacobiMatrix.GetRows;i++)
            {
                for(int j=0;j<JacobiMatrix.GetColumns;j++)
                {
                    if (i == j) //Диагональные элементы
                    {
                        JacobiMatrix.matrix[i, j] = diagConjugateNodesVoltage.matrix[i,j]*ConductivityMatrix.matrix[i,j];
                    }
                    else //Недиагональные элементы
                    {
                        JacobiMatrix.matrix[i, j] = diagConjugateNodesVoltage.matrix[i, i] * ConductivityMatrix.matrix[i, j];
                    }
                }
            }

            return JacobiMatrix;
        }

        public static MatrixComplex NodesVoltageInitialAprox(Node[] nodesArray, int baseNumber)
        {
            MatrixComplex InitialAprox = new MatrixComplex(nodesArray.Length - 1, 1);

            int k = 0;
            for(int i=0;i<nodesArray.Length; i++)
            {
                if (i == baseNumber) continue;
                InitialAprox.matrix[k, 0] = new Complex(nodesArray[i].NodeNominalVoltage, 0);
                k++;
            }
            return InitialAprox;
        }

        public static MatrixComplex SolveEquation(MatrixComplex NodesVoltage, MatrixComplex ConductivityMatrixC, MatrixComplex baseConductivityC, MatrixComplex LoadMatrixC, Complex baseVoltage, double Accuracy, out int step)
        {
            int r = NodesVoltage.GetRows;
            int c = NodesVoltage.GetColumns;
            MatrixComplex diagNodesVoltage = new MatrixComplex(r, r);                                               //Create diagonal nodes voltage matrix
            MatrixComplex ConjLoadMatrix = MatrixComplex.ConjugateMatrixElements(LoadMatrixC);                     //Create matrix with conjugated voltages
            MatrixComplex ConjNodesVoltage = MatrixComplex.ConjugateMatrixElements(NodesVoltage);
            MatrixComplex PowerMismatch = new MatrixComplex(r, c);
            MatrixComplex PowerMismatch2 = new MatrixComplex(r, c);


            bool AccuracyIsReached = false;
            double EuqlidianNorm = 0;
            double EuqlidianNormN_0;
            int counter = -1;

            step = 0;

            while (!AccuracyIsReached)
            {
                EuqlidianNormN_0=EuqlidianNorm;

                //Метод Зейделя
                for (int i=0; i<NodesVoltage.GetRows;i++)
                {
                    NodesVoltage.matrix[i, 0] = ConjLoadMatrix.matrix[i, 0] / ConjNodesVoltage.matrix[i, 0]-baseConductivityC.matrix[i,0]*baseVoltage;

                    for (int j=0;j<NodesVoltage.GetRows;j++)
                    {
                        if (j == i) continue;
                        NodesVoltage.matrix[i, 0] -= NodesVoltage.matrix[j, 0] * ConductivityMatrixC.matrix[i, j]; 
                    }

                    NodesVoltage.matrix[i, 0] = NodesVoltage.matrix[i, 0] / ConductivityMatrixC.matrix[i, i];
                }

                ConjNodesVoltage = MatrixComplex.ConjugateMatrixElements(NodesVoltage);

                step++;
                
                for (int i = 0; i < NodesVoltage.GetRows; i++)
                {
                    PowerMismatch.matrix[i, 0] = baseConductivityC.matrix[i, 0] * baseVoltage;
                    for (int j = 0; j < NodesVoltage.GetRows; j++)
                    {
                        PowerMismatch.matrix[i, 0] += ConductivityMatrixC.matrix[i,j] * NodesVoltage.matrix[j, 0];
                    }
                    PowerMismatch.matrix[i, 0] *= ConjNodesVoltage.matrix[i, 0];
                    PowerMismatch.matrix[i, 0] -= ConjLoadMatrix.matrix[i, 0];
                }

                EuqlidianNorm = Math.Abs(PowerMismatch.matrix[0, 0].Real);
                for (int i = 1; i < r; i++)
                    if (EuqlidianNorm < Math.Abs(PowerMismatch.matrix[i, 0].Real))
                        EuqlidianNorm = Math.Abs(PowerMismatch.matrix[i, 0].Real);

                if (EuqlidianNorm < Accuracy) AccuracyIsReached = true;
                if (EuqlidianNormN_0 < EuqlidianNorm) counter++;
                if (counter>50)
                {
                    Console.WriteLine("Режим расходится, необходимо введение в допустимую область");
                    return NodesVoltage;
                }
                if (step>=500)
                {
                    Console.WriteLine("Расчет завершился, достигнуто максимальное количество итераций");
                    return NodesVoltage;
                }
            }
            return NodesVoltage;
        }
    }
}
