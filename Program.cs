using System;
using OfficeOpenXml;
using System.IO;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.ExtendedProperties;
using System.Diagnostics;
using DocumentFormat.OpenXml.Bibliography;

namespace SteadyStateCalculation
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            Stopwatch swWholeProgram = new Stopwatch();
            Stopwatch swPartial = new Stopwatch();
            swWholeProgram.Start();

            string pathFile = @"Solved.xlsx";

            int numberNodes;
            int numberLines;
            int numberComutations = 0;
            int step = 0;
            int baseNumber;

            
            LinkWithExcel link = new LinkWithExcel(pathFile);
            link.CheckWorkSheets();

            //Ввод данных о узлах и ветвях
            link.GetNumberNodesAndLines(out numberNodes, out numberLines, ref numberComutations);
            Console.WriteLine("Число узлов: "+numberNodes);
            Console.WriteLine("Число ветвей: "+numberLines);

            Node[] nodesArray = new Node[numberNodes];
            Line[] linesArray = new Line[numberLines];
            
            link.FillNodesArray(ref nodesArray, out baseNumber);
            link.FillLinesArray(ref linesArray);

            //TODO 2: (3) -> Расчет неверный см. TODO 1
            Selection();

            while (true)
            {
                switch (Console.ReadKey().KeyChar)
                {
                    case '1':
                        UseSolverMatlab(numberNodes, step, numberComutations, baseNumber,
            linesArray, nodesArray, swWholeProgram, swPartial, link);
                        Selection();
                        break;

                    case '2':
                        UseSolver(numberNodes, step, numberComutations, baseNumber,
            linesArray, nodesArray, swWholeProgram, swPartial, link);
                        Selection();

                        break;
                    case '3':
                        UseSolverComplex(numberNodes, step, numberComutations, baseNumber,
            linesArray, nodesArray, swWholeProgram, swPartial, link);
                        Selection();
                        break;

                    case 'в':
                    case 'b':
                        return;

                    default:
                        break;
                }
            }

            void Selection()
            {
                Console.WriteLine("Нажмите соответствующую клавишу\n" +
                                "1 - расчет управляющих воздействий (необходим установленный MatLab)\n" +
                                "2 - расчет напряжений узлов в действительных числах\n" +
                                "3 - расчет напряжений узлов в комплексных числах (Расчет неверный)\n" +
                                "в - выход");
            }
        } //> Main

        public static void UseSolverMatlab(int numberNodes, int step, int numberComutations, int baseNumber, 
            Line[] linesArray, Node[] nodesArray, Stopwatch swWholeProgram, Stopwatch sw, LinkWithExcel link)
        {
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
            for (int i = 0; i < numberComutations + 1; i++)
            {
                if (step > 0)
                {
                    bool ComutationIsFound = false;
                    for (int j = currentLineOff + 1; j < linesArray.Length; j++)
                    {
                        if ((linesArray[j].LineCommutation == true) && (!ComutationIsFound))
                        {
                            ComutationIsFound = true;
                            if (currentLineOff >= 0) linesArray[currentLineOff].LineState = false;
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
                link.OutputResult(EstimatedVoltage, EstimatedAngle, dP, dQ, baseNumber, baseVoltage, nodesArray, linesArray, step, TanFi);
                step++;
            }

            swWholeProgram.Stop();
            Console.WriteLine($"Расчет {step - 1} управляющих воздействий выполнен за {swWholeProgram.ElapsedMilliseconds} мс");
            #endregion

            //Открытие папки с результатом
            Process PrFolder = new Process();
            ProcessStartInfo psi = new ProcessStartInfo();
            psi.CreateNoWindow = true;
            psi.WindowStyle = ProcessWindowStyle.Normal;
            psi.FileName = "explorer";
            psi.Arguments = @"/n, /select, " + link.GetPathFile;
            PrFolder.StartInfo = psi;
            PrFolder.Start();

            sw.Reset();
        }

        public static void UseSolver(int numberNodes, int step, int numberComutations, int baseNumber,
            Line[] linesArray, Node[] nodesArray, Stopwatch swWholeProgram, Stopwatch sw, LinkWithExcel link)
        {
            #region RealCalc

            Matrix G;
            Matrix B;
            Matrix Gb;
            Matrix Bb;
            Matrix ConductivityMatrix;
            Matrix baseConductivity;
            Matrix NodesVoltageReal = new Matrix(numberNodes - 1);
            Matrix NodesVoltageImage = new Matrix(numberNodes - 1);
            Matrix NodesVoltage = new Matrix((numberNodes - 1) * 2);
            Matrix LoadReal = new Matrix(numberNodes - 1);
            Matrix LoadImage = new Matrix(numberNodes - 1);
            Matrix LoadMatrix;
            

            sw.Start();

            Solver.GetConductivityMatrix(nodesArray, linesArray, numberNodes, baseNumber, out G, out B, out Gb, out Bb, out ConductivityMatrix, out baseConductivity);
            Solver.GetLoadMatrix(nodesArray, baseNumber, ref LoadReal, ref LoadImage);
            LoadMatrix = Solver.LoadMatrix(LoadReal, LoadImage);
            Solver.NodesVoltageInitialAprox(nodesArray, baseNumber, ref NodesVoltageReal, ref NodesVoltageImage);
            NodesVoltage = Solver.NodesVoltage(NodesVoltageReal, NodesVoltageImage);
            double baseVoltage = nodesArray[baseNumber].NodeNominalVoltage;

            sw.Stop();
            sw.Reset();

            Console.WriteLine("Вектор напряжений\n" + NodesVoltageReal + "\n" + NodesVoltageImage + "\n");
            Console.WriteLine("Напряжение базисного узла: " + baseVoltage + "\n");
            Console.WriteLine("Вектор мощностей:\n" + LoadReal + "\n" + LoadImage + "\n");
            Console.WriteLine("Параметры схемы загружены за " + sw.ElapsedMilliseconds + " мс\n"); sw.Reset();


            double Accuracy = 0;

            sw.Start();

            Matrix EstimatedNodesVoltage = Solver.SolveEquation(NodesVoltage, G, B, Gb, Bb, NodesVoltageReal, NodesVoltageImage, LoadMatrix, ConductivityMatrix, baseConductivity, baseVoltage, Accuracy, out step);

            sw.Stop();
            Console.WriteLine($"Итераций расчета: {step}\nВремя расчета: {sw.ElapsedMilliseconds} мс\nТочность: {Accuracy}\n\n" + EstimatedNodesVoltage);

            sw.Reset();

            #endregion
        }

        public static void UseSolverComplex(int numberNodes, int step, int numberComutations, int baseNumber,
            Line[] linesArray, Node[] nodesArray, Stopwatch swWholeProgram, Stopwatch sw, LinkWithExcel link)
        {
            #region ComplexCalc
            
            sw.Start();

            MatrixComplex baseConductivityC = new MatrixComplex(numberNodes - 1);
            MatrixComplex ConductivityMatrixC = SolverComplex.GetConductivityMatrix(nodesArray, linesArray, numberNodes, ref baseConductivityC, ref baseNumber);
            MatrixComplex LoadMatrixC = SolverComplex.GetLoadMatrix(nodesArray, baseNumber);
            MatrixComplex NodesVoltageC = SolverComplex.NodesVoltageInitialAprox(nodesArray, baseNumber);
            Complex baseVoltageC = new Complex(nodesArray[baseNumber].NodeNominalVoltage, 0);

            sw.Stop();

            Console.WriteLine("Матрица проводимостей:\n" + ConductivityMatrixC);
            Console.WriteLine("Вектор проводимостей базисного узла\n" + baseConductivityC);
            Console.WriteLine("Вектор напряжений\n" + NodesVoltageC);
            Console.WriteLine("Напряжение базисного узла: " + baseVoltageC + "\n");
            Console.WriteLine("Вектор мощностей:\n" + LoadMatrixC + "\n");
            Console.WriteLine("Параметры схемы загружены за " + sw.ElapsedMilliseconds + " мс\n"); sw.Reset();

            double AccuracyC = 0.1; //МВт
            int stepC;

            sw.Reset();
            sw.Start();

            MatrixComplex EstimatedNodesVoltageC = SolverComplex.SolveEquation(NodesVoltageC, ConductivityMatrixC, baseConductivityC, LoadMatrixC, baseVoltageC, AccuracyC, out stepC);

            sw.Stop();
            Console.WriteLine($"Итераций расчета: {stepC}\nВремя расчета: {sw.ElapsedMilliseconds} мс\nТочность: {AccuracyC} МВт\n\n" + EstimatedNodesVoltageC);
            Console.WriteLine(EstimatedNodesVoltageC.matrix[0, 0].Module + "<" + EstimatedNodesVoltageC.matrix[0, 0].Angle);
            Console.WriteLine(EstimatedNodesVoltageC.matrix[1, 0].Module + "<" + EstimatedNodesVoltageC.matrix[1, 0].Angle);

            #endregion
        }
    }//> Program
}
