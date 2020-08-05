using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml;
using System.IO;
using System.Diagnostics;
using System.Security.Cryptography;
using DocumentFormat.OpenXml.ExtendedProperties;

namespace SteadyStateCalculation
{
    class LinkWithExcel
    {
        FileInfo fi;
        string PathFile;
        public LinkWithExcel(string PathFile)
        {
            fi = new FileInfo(PathFile);
            this.PathFile = PathFile;
        }

        public void GetNumberNodesAndLines(out int numberNodes, out int numberLines, ref int numberComutations)
        {
            using (ExcelPackage package = new ExcelPackage(fi))
            {
                ExcelWorksheet nodes = package.Workbook.Worksheets["Nodes"];
                ExcelWorksheet lines = package.Workbook.Worksheets["Lines"];

                numberNodes = 0;
                numberLines = 0;
                int i = 2;
                while (nodes.Cells[i, 1].Value !=null)
                {
                    i++;
                    numberNodes++;
                }
                i = 2;
                while (lines.Cells[i, 1].Value!=null)
                {
                    if (lines.Cells[i, 1].Value.ToString() == "1") numberComutations++;
                    i++;
                    numberLines++;
                }
            }
        }

        public void CheckResultWorkSheet(Node[] nodesArray, Line[] linesArray, int baseNumber)
        {
            using (ExcelPackage package = new ExcelPackage(fi))
            {
                #region CheckResult
                if (package.Workbook.Worksheets["Result"] == null)
                {
                    package.Workbook.Worksheets.Add("Result");
                }
                ExcelWorksheet result = package.Workbook.Worksheets["Result"];
                ExcelWorksheet nodes = package.Workbook.Worksheets["Nodes"];
                ExcelWorksheet lines = package.Workbook.Worksheets["Lines"];

                result.Cells[1, 1, 1, nodesArray.Length + 2].Merge = true;
                result.Cells[1, 1].Value = "Место приложения управляющего воздействия";
                result.Cells[2, 1, 3, 1].Merge = true;
                result.Cells[2, 1].Value = "Отключение";
                result.Cells[2, 2, 3, 2].Merge = true;
                result.Cells[2, 2].Value = "Номер начала";
                result.Cells[2, 3, 3, 3].Merge = true;
                result.Cells[2, 3].Value = "Номер конца";

                result.Cells[4, 1].Value = "Нормальный режим";
                int j = 4;
                for (int i = 0; i < nodesArray.Length; i++)
                {
                    if (i == baseNumber) continue;
                    else
                    {
                        result.Cells[2, j].Value = nodesArray[i].NodeName;
                        result.Cells[3, j++].Value = nodesArray[i].NodeNumber;
                    }
                }
                j = 5;
                for (int i = 0; i < linesArray.Length; i++)
                {
                    if (linesArray[i].LineCommutation == true)
                    {
                        result.Cells[j, 1].Value = linesArray[i].LineName;
                        result.Cells[j, 2].Value = linesArray[i].LineStart;
                        result.Cells[j++, 3].Value = linesArray[i].LineEnd;
                    }
                }
                #endregion

                package.Save();
            }

        }

        public void FillNodesArray(ref Node[] nodesArray, out int baseNumber)
        {
            baseNumber = -1;
            using (ExcelPackage package = new ExcelPackage(fi))
            {
                ExcelWorksheet nodes = package.Workbook.Worksheets["Nodes"];
                for (int j = 0, i =2; j < nodesArray.Length; j++, i++)
                {
                    nodesArray[j] = new Node();
                    
                    nodesArray[j].NodeLoad = new Complex();
                    nodesArray[j].NodeGeneration = new Complex();
                    nodesArray[j].NodeConductivity = new Complex();
                    
                    nodesArray[j].NodeState = Boolean.Parse(nodes.Cells[i,1].Value.ToString());
                    nodesArray[j].NodeType = nodes.Cells[i,2].Value.ToString();
                    if (nodesArray[j].NodeType == "База") baseNumber = j;
                    nodesArray[j].NodeNumber = Int32.Parse(nodes.Cells[i,3].Value.ToString());
                    nodesArray[j].NodeName = nodes.Cells[i,4].Value.ToString();
                    nodesArray[j].NodeNominalVoltage = Double.Parse(nodes.Cells[i,5].Value.ToString());
                    nodesArray[j].NodeLoad.Real = Double.Parse(nodes.Cells[i,8].Value.ToString());
                    nodesArray[j].NodeLoad.Image = Double.Parse(nodes.Cells[i,9].Value.ToString());
                    nodesArray[j].NodeGeneration.Real = Double.Parse(nodes.Cells[i,10].Value.ToString());
                    nodesArray[j].NodeGeneration.Image = Double.Parse(nodes.Cells[i,11].Value.ToString());
                    nodesArray[j].NodeConductivity.Real = Double.Parse(nodes.Cells[i, 15].Value.ToString()) / 1000000.0;
                    nodesArray[j].NodeConductivity.Image = (-1.0)*Double.Parse(nodes.Cells[i,16].Value.ToString())/1000000.0;
                }
            }
        }

        public void FillLinesArray(ref Line[] linesArray)
        {
            using (ExcelPackage package = new ExcelPackage(fi))
            {
                ExcelWorksheet lines = package.Workbook.Worksheets["Lines"];
                for (int j = 0, i = 2; j < linesArray.Length; j++, i++)
                {
                    linesArray[j] = new Line();
                    
                    linesArray[j].LineResistance = new Complex();
                    linesArray[j].LineConductivity = new Complex();
                    linesArray[j].LineLoadStart = new Complex();
                    linesArray[j].LineLoadEnd = new Complex();
                    linesArray[j].LineLosses = new Complex();
                    string temp = lines.Cells[i, 1].Value.ToString();
                    if ((temp == "0")||(temp=="")||(temp==null)) linesArray[j].LineCommutation = false;
                    else linesArray[j].LineCommutation = true;
                    if (lines.Cells[i, 2].Value.ToString() == "0") linesArray[j].LineState = false;
                    else linesArray[j].LineState = true;
                    linesArray[j].LineStart = Int32.Parse(lines.Cells[i, 4].Value.ToString());
                    linesArray[j].LineEnd = Int32.Parse(lines.Cells[i, 5].Value.ToString());
                    linesArray[j].LineName = lines.Cells[i, 8].Value.ToString();
                    linesArray[j].LineResistance.Real = Double.Parse(lines.Cells[i, 9].Value.ToString());
                    linesArray[j].LineResistance.Image = Double.Parse(lines.Cells[i, 10].Value.ToString());
                    linesArray[j].LineConductivity.Real = Double.Parse(lines.Cells[i, 11].Value.ToString()) / 1000000.0 / 2.0; //Проводимость, приходящаяся на 1 узел
                    linesArray[j].LineConductivity.Image = (-1.0)*Double.Parse(lines.Cells[i, 12].Value.ToString())/1000000.0/2.0; //Проводимость, приходящаяся на 1 узел
                }
            }
        }

        public void OutputResult(Matrix EstimatedVoltage, Matrix EstimatedAngle, Matrix dP, Matrix dQ, double baseNumber, double baseVoltage, Node[] nodesArray, Line[] linesArray, int step, double cosfi, Matrix TanFi)
        {
            using (ExcelPackage package = new ExcelPackage(fi))
            {
                if (package.Workbook.Worksheets["Result"]==null)
                {
                    package.Workbook.Worksheets.Add("Result");
                }
                ExcelWorksheet result = package.Workbook.Worksheets["Result"];
                ExcelWorksheet nodes = package.Workbook.Worksheets["Nodes"];
                ExcelWorksheet lines = package.Workbook.Worksheets["Lines"];

                int j = 0;
                for(int i=0; i<EstimatedVoltage.GetRows+1;i++)
                {
                    if (i == baseNumber)
                    {
                        nodes.Cells[i + 2, 17].Value = baseVoltage;
                        nodes.Cells[i + 2, 18].Value = 0;
                    }
                    else
                    {
                        nodes.Cells[i + 2, 17].Value = EstimatedVoltage.matrix[j,0];
                        nodes.Cells[i + 2, 18].Value = EstimatedAngle.matrix[j++,0];
                    }
                }

                //Вывод невязок режима
                for (int i = 0; i < dP.GetRows; i++)
                {
                    ////Вывод с учетом реактивной мощности
                    double temp = (-1.0) * dQ.matrix[i, 0] / TanFi.matrix[i, 0];
                    if (Math.Abs(temp) > Math.Abs(dP.matrix[i, 0]))
                    {
                        if (temp > 0)
                        {
                            result.Cells[4 + step, i + 4].Value = 0;
                            result.Cells[4 + step, i + 4 + 13].Value = 0;
                        }
                        else
                        {
                            result.Cells[4 + step, i + 4].Value = temp;
                            result.Cells[4 + step, i + 4 + 13].Value = -dQ.matrix[i, 0];
                        }
                        //if (temp > -1.0) result.Cells[4 + step, i + 4].Value = "0";
                        //else result.Cells[4 + step, i + 4].Value = temp;
                    }
                    else
                    {
                        if (dP.matrix[i, 0] > 0)
                        {
                            result.Cells[4 + step, i + 4].Value = 0;
                            result.Cells[4 + step, i + 4 + 13].Value = 0;
                        }
                        else
                        {
                            result.Cells[4 + step, i + 4].Value = dP.matrix[i, 0];
                            result.Cells[4 + step, i + 4 + 13].Value = dP.matrix[i, 0] * TanFi.matrix[i, 0];
                        }

                        //if (dP.matrix[i, 0] > -1.0) result.Cells[4 + step, i + 4].Value = "0";
                        //else result.Cells[4 + step, i + 4].Value = dP.matrix[i, 0];
                    }


                    ////Вывод просто невязок
                    //if (dP.matrix[i, 0] == 0) result.Cells[4 + step, i + 4].Value = "0";
                    //else result.Cells[4 + step, i + 4].Value = dP.matrix[i, 0];

                    //if (-dQ.matrix[i, 0] == 0) result.Cells[4 + step, i + 4 + 13].Value = "0";
                    //else result.Cells[4 + step, i + 4 + 13].Value = -dQ.matrix[i, 0];




                    //if (Math.Abs(dP.matrix[i, 0]) < 2) result.Cells[4 + step, i + 4].Value = "0";
                    //else
                    //{
                    //    double temp = (-1.0) * dQ.matrix[i, 0] / Math.Tan(Math.Acos(cosfi));
                    //    if (Math.Abs(temp) > Math.Abs(dP.matrix[i, 0])) result.Cells[4 + step, i + 4].Value = $"{temp:0}";
                    //    else result.Cells[4 + step, i + 4].Value = $"{dP.matrix[i, 0]:0}";
                    //}
                }
                step++;

                package.Save();
            }

        }
    }
}
