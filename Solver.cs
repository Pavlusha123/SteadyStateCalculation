using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Math;
using System;
using System.Collections.Generic;
using System.Text;

namespace SteadyStateCalculation
{
    class Solver
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
            Console.WriteLine(ConductivityMatrix);
            Console.WriteLine();
            Console.WriteLine(baseConductivity);
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

        public static void GetLoadMatrix(Node[] nodesArray, int baseNumber, ref Matrix LoadReal, ref Matrix LoadImage)
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

        public static Matrix CalcVectorNevyazok(Matrix NodesVoltage, Matrix ConductivityMatrix)
        {
            return (Matrix.CreateDiagMatrix(NodesVoltage) * ConductivityMatrix * NodesVoltage);
        }

        public static void NormaNevyazok(Matrix VectorNevyazok, out double NevyazkaP, out double NevyazkaQ)
        {
            NevyazkaP = 0;
            NevyazkaQ = 0;
            for(int i=0;i<VectorNevyazok.GetRows/2;i++)
            {
                NevyazkaP += (VectorNevyazok.matrix[i, 0] * VectorNevyazok.matrix[i, 0]);
            }
            for (int i = VectorNevyazok.GetRows / 2; i < VectorNevyazok.GetRows; i++)
            {
                NevyazkaQ += (VectorNevyazok.matrix[i, 0] * VectorNevyazok.matrix[i, 0]);
            }
            NevyazkaP = Math.Pow(NevyazkaP, 0.5);
            NevyazkaQ = Math.Pow(NevyazkaQ, 0.5);
        }
        
        //TODO: Исправить матрицу Якоби
        public static Matrix CalcJacobiMatrix(Matrix G,Matrix B,Matrix Gb,Matrix Bb,Matrix NodesVoltageReal,Matrix NodesVoltageImage,double baseVoltage)
        {
            Matrix JacobiMatrix = new Matrix(G.GetRows << 1, G.GetColumns << 1);

            //Form second quadrant
            for(int i=0;i<G.GetRows;i++)
            {
                for(int j=0;j<G.GetColumns;j++)
                {
                    JacobiMatrix.matrix[i, j] = 0;
                }
            }

            //Form third quadrant
            int i1 = 0;
            for (int i = G.GetRows; i < JacobiMatrix.GetRows; i++)
            {
                for (int j = 0; j < G.GetColumns; j++)
                {
                    if (i1 == j) //Диагональные элементы
                    {
                        for (int k = 0; k < B.GetColumns; k++)
                        {
                            JacobiMatrix.matrix[i, j] += B.matrix[i1, k] * NodesVoltageReal.matrix[k, 0];
                        }
                        JacobiMatrix.matrix[i, j] += B.matrix[i1, i1] * NodesVoltageReal.matrix[i1, 0] + Bb.matrix[i1, 0] * baseVoltage;
                    } //Недиагональные элементы
                    else JacobiMatrix.matrix[i, j] = B.matrix[i1, j] * NodesVoltageReal.matrix[i1, 0];
                }
                i1++;
            }


            //Form first quadrant
            int j1 = 0;
            for (int i = 0; i < B.GetRows; i++)
            {
                j1 = 0;
                for (int j = B.GetColumns; j < JacobiMatrix.GetColumns; j++)
                {
                    if(i==j1) //Диагональные элементы
                    {
                        for (int k = 0; k < B.GetColumns; k++)
                        {
                            JacobiMatrix.matrix[i, j] += (-1.0)*B.matrix[i,k]*NodesVoltageReal.matrix[k,0];
                        }
                        JacobiMatrix.matrix[i, j] += B.matrix[i,i]*NodesVoltageReal.matrix[i,0] - Bb.matrix[i,0]*baseVoltage;
                    } //Недиагональные элементы
                    else JacobiMatrix.matrix[i, j] = B.matrix[i,j1]*NodesVoltageReal.matrix[i,0];
                    j1++;
                }
            }

            //Form fourth quadrant
            for (int i = B.GetRows; i < JacobiMatrix.GetRows; i++, i1++)
            {
                for (int j = B.GetColumns; j < JacobiMatrix.GetColumns; j++)
                {
                    JacobiMatrix.matrix[i, j] = 0;
                }
            }

            return JacobiMatrix;
        }

        public static void NodesVoltageInitialAprox(Node[] nodesArray, int baseNumber, ref Matrix NodesVoltageReal, ref Matrix NodesVoltageImage)
        {
            int k = 0;
            for (int i = 0; i < nodesArray.Length; i++)
            {
                if (i == baseNumber) continue;
                NodesVoltageReal.matrix[k, 0] = nodesArray[i].NodeNominalVoltage;
                NodesVoltageImage.matrix[k, 0] = 0;
                k++;
            }

        }

        public static void NodesVoltageRealImage(Matrix NodesVoltage, ref Matrix NodesVoltageReal, ref Matrix NodesVoltageImage)
        {
            int k = 0;
            for(int i=0;i<NodesVoltage.GetRows/2;i++)
            {
                NodesVoltageReal.matrix[i, 0] = NodesVoltage.matrix[i, 0];
            }
            for(int i=NodesVoltage.GetRows/2;i<NodesVoltage.GetRows;i++)
            {
                NodesVoltageImage.matrix[k++, 0] = NodesVoltage.matrix[i, 0];
            }
        }

        public static Matrix NodesVoltage(Matrix NodesVoltageReal, Matrix NodesVoltageImage)
        {
            Matrix NodesVoltage = new Matrix(NodesVoltageReal.GetRows * 2, 1);

            int j = 0;
            for(int i=0;i<NodesVoltage.GetRows;i++)
            {
                if (i < NodesVoltage.GetRows / 2)
                    NodesVoltage.matrix[i, 0] = NodesVoltageReal.matrix[i, 0];
                else NodesVoltage.matrix[i, 0] = NodesVoltageImage.matrix[j++, 0];
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

        public static Matrix SolveEquation(Matrix NodesVoltage,Matrix G,Matrix B,Matrix Gb,Matrix Bb,Matrix NodesVoltageReal,Matrix NodesVoltageImage,Matrix LoadMatrix,Matrix ConductivityMatrix,Matrix baseConductivity,double baseVoltage,double Accuracy, out int step)
        {
            int r = NodesVoltage.GetRows;
            int c = NodesVoltage.GetColumns;
            Matrix NodesVoltageN_0 = new Matrix(r, c);                //Create N-1 step nodes voltage
            Matrix PowerMismatch = new Matrix(r, c);
            Matrix JacobiMatrix;
            Matrix InvertedJacobiMatrix;
            Matrix dU;

            bool AccuracyIsReached = false;
            double EuqlidianNorm = 0;
            double EuqlidianNormN_0;
            int counter = -1;

            step = 0;

            while (!AccuracyIsReached)
            {
                EuqlidianNormN_0 = EuqlidianNorm;

                NodesVoltageN_0 = Matrix.CopyMatrix(NodesVoltage);                                                         //Save N-1 step nodes voltage

                JacobiMatrix = Solver.CalcJacobiMatrix(G, B, Gb, Bb, NodesVoltageReal, NodesVoltageImage, baseVoltage);

                InvertedJacobiMatrix = Matrix.CreateInvertedMatrixGauss(JacobiMatrix);

                PowerMismatch = Matrix.CreateDiagMatrix(Matrix.ConjugateMatrixElements(NodesVoltage)) * ConductivityMatrix * NodesVoltage + Matrix.CreateDiagMatrix(Matrix.ConjugateMatrixElements(NodesVoltage)) * baseConductivity * baseVoltage;

                dU = InvertedJacobiMatrix * (PowerMismatch);

                NodesVoltage -= dU;

                Solver.NodesVoltageRealImage(NodesVoltage, ref NodesVoltageReal, ref NodesVoltageImage);

                step++;

                //PowerMismatch = JacobiMatrix * NodesVoltage;
                EuqlidianNorm = PowerMismatch.matrix[0, 0];
                for (int i = 1; i < PowerMismatch.GetRows; i++)
                    if (EuqlidianNorm < PowerMismatch.matrix[i, 0])
                        EuqlidianNorm = PowerMismatch.matrix[i, 0];

                if (EuqlidianNorm < Accuracy) AccuracyIsReached = true;
                if (EuqlidianNormN_0 < EuqlidianNorm) counter++;
                if ((counter > 3)|(step>20))
                {
                    Console.WriteLine("Режим расходится, необходимо введение в допустимую область");
                    return NodesVoltage;
                }

            }
            return NodesVoltage;

        }

    }
}
