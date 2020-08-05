using System;
using OfficeOpenXml;
using System.IO;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.ExtendedProperties;

namespace SteadyStateCalculation
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            System.Diagnostics.Stopwatch swWholePorgram = new System.Diagnostics.Stopwatch();
            swWholePorgram.Start();

            string pathFile = @"Solved.xlsx";
            double cosfi = 0.9; //Задание коэффициента нагрузки
            int maxiter = 1; //Кол-во итераций расчета управляющих воздействий

            int numberNodes;
            int numberLines;
            int numberComutations = 0;
            int step = 0;
            int baseNumber;

            //Ввод данных о узлах и ветвях
            LinkWithExcel link = new LinkWithExcel(pathFile);
            link.GetNumberNodesAndLines(out numberNodes, out numberLines, ref numberComutations);
            Console.WriteLine("Число узлов: "+numberNodes);
            Console.WriteLine("Число ветвей: "+numberLines);

            Node[] nodesArray = new Node[numberNodes];
            Line[] linesArray = new Line[numberLines];
            
            link.FillNodesArray(ref nodesArray, out baseNumber);
            link.FillLinesArray(ref linesArray);
            link.CheckResultWorkSheet(nodesArray, linesArray, baseNumber);

            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
            sw.Start();
            sw.Stop();
            sw.Reset();

            #region MatLabOptimization
            //Создание массивов данных
            Matrix G;
            Matrix B;
            Matrix Gb;
            Matrix Bb;
            Matrix ConductivityMatrix;
            Matrix baseConductivity;
            Matrix NodesVoltageModule = new Matrix(numberNodes - 1);
            Matrix NodesVoltageAngle = new Matrix(numberNodes - 1);
            Matrix NodesVoltage = new Matrix((numberNodes - 1) * 2);
            Matrix LoadMatrixReal = new Matrix(numberNodes - 1);
            Matrix LoadMatrixImage = new Matrix(numberNodes - 1);
            Matrix TanFi = new Matrix(numberNodes - 1);
            Matrix LoadMatrix;
            Matrix dP;
            Matrix dQ;


            //Продолжать расчет, пока не рассмотрены все случаи
            int currentLineOff = -1;
            for (int i = 0; i < numberComutations+1; i++)
            {
                if(step>0)
                {
                    bool ComutationIsFound = false;
                    for(int j=currentLineOff+1;j<linesArray.Length;j++)
                    {
                        if((linesArray[j].LineCommutation==true)&&(!ComutationIsFound))
                        {
                            ComutationIsFound = true;
                            if(currentLineOff>=0) linesArray[currentLineOff].LineState = false;
                            linesArray[j].LineState = true;
                            currentLineOff = j;
                        }
                    }
                    Console.WriteLine($"\n\nОтключена линия {linesArray[currentLineOff].LineName}");
                }
                else Console.WriteLine($"\n\nНормальный режим");

                //Составление массивов данных о сети на основе информации о узлах и ветвях
                sw.Start();

                SolverMatLab.GetConductivityMatrix(nodesArray, linesArray, numberNodes, baseNumber, out G, out B, out Gb, out Bb, out ConductivityMatrix, out baseConductivity);
                SolverMatLab.GetLoadMatrix(nodesArray, baseNumber, ref LoadMatrixReal, ref LoadMatrixImage, ref TanFi);
                LoadMatrix = SolverMatLab.LoadMatrix(LoadMatrixReal, LoadMatrixImage);
                SolverMatLab.NodesVoltageInitialAprox(nodesArray, baseNumber, ref NodesVoltageModule, ref NodesVoltageAngle);
                NodesVoltage = SolverMatLab.NodesVoltage(NodesVoltageModule, NodesVoltageAngle);
                double baseVoltage = nodesArray[baseNumber].NodeNominalVoltage;

                sw.Stop();

                for (int iter = 0; iter < maxiter; iter++)
                {
                    if (step == 0) iter = maxiter - 1;
                    //Вывод данных в консоль
                    //Console.WriteLine("Матрица G\n" + G);
                    //Console.WriteLine("Матрица B\n" + B);
                    //Console.WriteLine("Матрица Gb\n" + Gb);
                    //Console.WriteLine("Матрица Bb\n" + Bb);
                    //Console.WriteLine("Вектор напряжений\n" + NodesVoltage);
                    //Console.WriteLine("Напряжение базисного узла: " + baseVoltage + "\n");
                    //Console.WriteLine("Вектор мощностей:\n" + LoadMatrixReal + "\n" + LoadMatrixImage + "\n");
                    Console.WriteLine("Параметры схемы загружены за " + sw.ElapsedMilliseconds + " мс\n");
                    sw.Reset();

                    //Оптимизационный метод решения установившегося режима в Matlab
                    Matrix EstimatedVoltage;
                    Matrix EstimatedAngle;
                    sw.Start();
                    SolverMatLab.SolveEquation(NodesVoltage, G, B, Gb, Bb, NodesVoltageModule, NodesVoltageAngle, LoadMatrixReal, LoadMatrixImage, ConductivityMatrix, baseConductivity, baseVoltage, out EstimatedVoltage, out EstimatedAngle, out dP, out dQ);
                    Console.WriteLine($"Решение получено за {sw.ElapsedMilliseconds} мс");
                    sw.Stop();
                    sw.Reset();

                    //Вывод результатов расчета
                    link.OutputResult(EstimatedVoltage, EstimatedAngle, dP, dQ, baseNumber, baseVoltage, nodesArray, linesArray, step, cosfi, TanFi);
                    step++;

                    ////Изменение нагрузок после итерации
                    //for (int i1 = 0; i1 < LoadMatrixReal.GetRows; i1++)
                    //{
                    //    double temp = (-1.0) * dQ.matrix[i, 0] / TanFi.matrix[i, 0];
                    //    if (Math.Abs(temp) > Math.Abs(dP.matrix[i, 0]))
                    //    {
                    //        if (temp > 0)
                    //        {
                    //            LoadMatrixReal.matrix[i1, 0] += 0;
                    //            LoadMatrixImage.matrix[i1, 0] += 0;
                    //        }
                    //        else
                    //        {
                    //            LoadMatrixReal.matrix[i1, 0] += temp;
                    //            LoadMatrixImage.matrix[i1, 0] += -dQ.matrix[i, 0];
                    //        }
                    //    }
                    //    else
                    //    {
                    //        if (dP.matrix[i, 0] > 0)
                    //        {
                    //            LoadMatrixReal.matrix[i1, 0] += 0;
                    //            LoadMatrixImage.matrix[i1, 0] += 0;
                    //        }
                    //        else
                    //        {
                    //            LoadMatrixReal.matrix[i1, 0] += dP.matrix[i, 0];
                    //            LoadMatrixImage.matrix[i1, 0] += dP.matrix[i, 0] * TanFi.matrix[i, 0];
                    //        }
                    //    }
                    //}
                    #endregion
                }
            }
            swWholePorgram.Stop();
            Console.WriteLine($"Расчет {step-1} управляющих воздействий выполнен за {swWholePorgram.ElapsedMilliseconds} мс");
            #region RealCalc
            //Matrix G;
            //Matrix B;
            //Matrix Gb;
            //Matrix Bb;
            //Matrix ConductivityMatrix;
            //Matrix baseConductivity;
            //Matrix NodesVoltageReal = new Matrix(numberNodes - 1);
            //Matrix NodesVoltageImage = new Matrix(numberNodes - 1);
            //Matrix NodesVoltage = new Matrix((numberNodes - 1) * 2);
            //Matrix LoadReal = new Matrix(numberNodes - 1);
            //Matrix LoadImage = new Matrix(numberNodes - 1);
            //Matrix LoadMatrix;
            //int step;

            //sw.Start();

            //Solver.GetConductivityMatrix(nodesArray, linesArray, numberNodes, baseNumber, out G, out B, out Gb, out Bb, out ConductivityMatrix, out baseConductivity);
            //Solver.GetLoadMatrix(nodesArray, baseNumber, ref LoadReal, ref LoadImage);
            //LoadMatrix = Solver.LoadMatrix(LoadReal, LoadImage);
            //Solver.NodesVoltageInitialAprox(nodesArray, baseNumber, ref NodesVoltageReal, ref NodesVoltageImage);
            //NodesVoltage = Solver.NodesVoltage(NodesVoltageReal, NodesVoltageImage);
            //double baseVoltage = nodesArray[baseNumber].NodeNominalVoltage;

            //sw.Stop();
            //sw.Reset();

            //Console.WriteLine("Вектор напряжений\n" + NodesVoltageReal + "\n" + NodesVoltageImage + "\n");
            //Console.WriteLine("Напряжение базисного узла: " + baseVoltage + "\n");
            //Console.WriteLine("Вектор мощностей:\n" + LoadReal + "\n" + LoadImage + "\n");
            //Console.WriteLine("Параметры схемы загружены за " + sw.ElapsedMilliseconds + " мс\n"); sw.Reset();


            //double Accuracy = 0;

            //sw.Start();

            //Matrix EstimatedNodesVoltage = Solver.SolveEquation(NodesVoltage, G, B, Gb, Bb, NodesVoltageReal, NodesVoltageImage, LoadMatrix, ConductivityMatrix, baseConductivity, baseVoltage, Accuracy, out step);

            //sw.Stop();
            //Console.WriteLine($"Итераций расчета: {step}\nВремя расчета: {sw.ElapsedMilliseconds} мс\nТочность: {Accuracy}\n\n" + EstimatedNodesVoltage);

            //sw.Reset();

            #region MatrixTest
            //Matrix m = new Matrix(3, 3);
            //Random rnd = new Random();
            //for (int i = 0; i < m.GetRows; i++)
            //    for (int j = 0; j < m.GetColumns; j++)
            //        m.matrix[i, j] = rnd.Next(-2,2);
            //Console.WriteLine(m);
            //Matrix p = Matrix.CreateUnitMatrix(6);
            //Matrix.SwapRows(m, 0, 5);
            //Matrix.SwapRows(p, 0, 5);
            //Console.WriteLine(m);
            //Console.WriteLine();
            //Console.WriteLine(p);
            //Console.WriteLine();
            //Console.WriteLine(p*m);

            //m.matrix[0, 0] = 0;
            //m.matrix[0, 1] = 1;
            //m.matrix[0, 2] = 0;
            //m.matrix[1, 0] = -1;
            //m.matrix[1, 1] = 0;
            //m.matrix[1, 2] = 1;
            //m.matrix[2, 0] = -2;
            //m.matrix[2, 1] = -1;
            //m.matrix[2, 2] = -1;
            //Console.WriteLine(m);

            //sw.Start();
            //Console.WriteLine(Matrix.CreateInvertedMatrixGauss(m));
            //sw.Stop();
            //Console.WriteLine("\n");
            //Console.WriteLine(sw.ElapsedTicks);
            //Console.WriteLine("\n\n");

            //sw.Reset();
            //sw.Start();
            //Console.WriteLine(Matrix.CreateInvertedMatrix(m));
            //sw.Stop();
            //Console.WriteLine("\n");
            //Console.WriteLine(sw.ElapsedTicks);
            //sw.Reset();
            #endregion
            #endregion

            #region ComplexCalc

            //sw.Start();

            //MatrixComplex baseConductivityC = new MatrixComplex(numberNodes - 1);
            //MatrixComplex ConductivityMatrixC = SolverComplex.GetConductivityMatrix(nodesArray, linesArray, numberNodes, ref baseConductivityC, ref baseNumber);
            //MatrixComplex LoadMatrixC = SolverComplex.GetLoadMatrix(nodesArray, baseNumber);
            //MatrixComplex NodesVoltageC = SolverComplex.NodesVoltageInitialAprox(nodesArray, baseNumber);
            //Complex baseVoltageC = new Complex(nodesArray[baseNumber].NodeNominalVoltage, 0);

            //sw.Stop();

            //Console.WriteLine("Матрица проводимостей:\n" + ConductivityMatrixC);
            //Console.WriteLine("Вектор проводимостей базисного узла\n" + baseConductivityC);
            //Console.WriteLine("Вектор напряжений\n" + NodesVoltageC);
            //Console.WriteLine("Напряжение базисного узла: " + baseVoltageC + "\n");
            //Console.WriteLine("Вектор мощностей:\n" + LoadMatrixC + "\n");
            //Console.WriteLine("Параметры схемы загружены за " + sw.ElapsedMilliseconds + " мс\n"); sw.Reset();

            //double AccuracyC = 0.1; //МВт
            //int stepC;

            //sw.Reset();
            //sw.Start();

            //MatrixComplex EstimatedNodesVoltageC = SolverComplex.SolveEquation(NodesVoltageC, ConductivityMatrixC, baseConductivityC, LoadMatrixC, baseVoltageC, AccuracyC, out stepC);

            //sw.Stop();
            //Console.WriteLine($"Итераций расчета: {stepC}\nВремя расчета: {sw.ElapsedMilliseconds} мс\nТочность: {AccuracyC} МВт\n\n" + EstimatedNodesVoltageC);
            //Console.WriteLine(EstimatedNodesVoltageC.matrix[0, 0].Module + "<" + EstimatedNodesVoltageC.matrix[0, 0].Angle);
            //Console.WriteLine(EstimatedNodesVoltageC.matrix[1, 0].Module + "<" + EstimatedNodesVoltageC.matrix[1, 0].Angle);

            #region MatrixTest
            //MatrixComplex m = new MatrixComplex(3, 3);
            //m.matrix[0, 0] = new Complex(0,0);
            //m.matrix[0, 1] = new Complex(1, 2);
            //m.matrix[0, 2] = new Complex(4,3);
            //m.matrix[1, 0] = new Complex(5,8);
            //m.matrix[1, 1] = new Complex(1,1);
            //m.matrix[1, 2] = new Complex(3,1);
            //m.matrix[2, 0] = new Complex(4,2);
            //m.matrix[2, 1] = new Complex(3,5);
            //m.matrix[2, 2] = new Complex(2,0);
            //Console.WriteLine(m);
            //Console.WriteLine();
            //Console.WriteLine(MatrixComplex.CreateInvertedMatrixGauss(m));

            //MatrixComplex NodesVoltage = Solver.NodesVoltageInitialAprox(nodesArray, baseNumber);

            //Complex baseVoltage = new Complex(nodesArray[baseNumber].NodeNominalVoltage, 0);

            //MatrixComplex JacobiMatrix = Solver.CalcJacobiMatrix(ConductivityMatrix, baseConductivity, NodesVoltage, baseVoltage);

            //Console.WriteLine(JacobiMatrix);
            //Console.WriteLine();
            //Console.WriteLine(ConductivityMatrix);
            //Console.WriteLine();
            //Console.WriteLine(MatrixComplex.CreateInvertedMatrixGauss(JacobiMatrix));
            //Console.WriteLine((MatrixComplex.CreateInvertedMatrixGauss(JacobiMatrix) * MatrixComplex.ConjugateMatrixElements(LoadMatrix)));


            //NodesVoltage = NodesVoltage + (MatrixComplex.CreateInvertedMatrixGauss(JacobiMatrix) * LoadMatrix);
            //Console.WriteLine(NodesVoltage);
            #endregion
            #endregion
        }
    }
}
