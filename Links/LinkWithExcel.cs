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

        public string GetPathFile 
        { 
            get { return PathFile; } 
            private set { PathFile = value; } 
        }

        public LinkWithExcel(string PathFile)
        {
            fi = new FileInfo(PathFile);
            this.GetPathFile = PathFile;
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

        public void CheckWorkSheets()
        {
            using (ExcelPackage package = new ExcelPackage(fi))
            {
                
                #region CheckNodes
                if (package.Workbook.Worksheets["Nodes"] == null)
                {
                    package.Workbook.Worksheets.Add("Nodes");
                }
                ExcelWorksheet nodes = package.Workbook.Worksheets["Nodes"];
                
                nodes.Cells[1, 1].Value = "Отключен";
                nodes.Cells[1, 2].Value = "Тип";
                nodes.Cells[1, 2].Value = "Номер";
                nodes.Cells[1, 4].Value = "Название";
                nodes.Cells[1, 5].Value = "U_ном";
                nodes.Cells[1, 6].Value = "N_схн";
                nodes.Cells[1, 7].Value = "Район";
                nodes.Cells[1, 8].Value = "P_н";
                nodes.Cells[1, 9].Value = "Q_н";
                nodes.Cells[1, 10].Value = "P_г";
                nodes.Cells[1, 11].Value = "Q_г";
                nodes.Cells[1, 12].Value = "V_зд";
                nodes.Cells[1, 13].Value = "Q_min";
                nodes.Cells[1, 14].Value = "Q_max";
                nodes.Cells[1, 15].Value = "G_ш";
                nodes.Cells[1, 16].Value = "B_ш";
                nodes.Cells[1, 17].Value = "V";
                nodes.Cells[1, 18].Value = "Delta";
                #endregion

                #region CheckLines
                if (package.Workbook.Worksheets["Lines"] == null)
                {
                    package.Workbook.Worksheets.Add("Lines");
                }
                ExcelWorksheet lines = package.Workbook.Worksheets["Lines"];

                lines.Cells[1, 1].Value = "Отключение";
                lines.Cells[1, 2].Value = "Состояние";
                lines.Cells[1, 3].Value = "Тип";
                lines.Cells[1, 4].Value = "N_нач";
                lines.Cells[1, 5].Value = "N_кон";
                lines.Cells[1, 6].Value = "N_п";
                lines.Cells[1, 7].Value = "ID Группы";
                lines.Cells[1, 8].Value = "Название";
                lines.Cells[1, 9].Value = "R";
                lines.Cells[1, 10].Value = "X";
                lines.Cells[1, 11].Value = "G";
                lines.Cells[1, 12].Value = "B";
                lines.Cells[1, 13].Value = "Ктр";
                lines.Cells[1, 14].Value = "N_анц";
                lines.Cells[1, 15].Value = "БД_анц";
                lines.Cells[1, 16].Value = "P_нач";
                lines.Cells[1, 17].Value = "Q_нач";
                lines.Cells[1, 18].Value = "dP";
                lines.Cells[1, 19].Value = "dQ";
                lines.Cells[1, 20].Value = "P_кон";
                lines.Cells[1, 21].Value = "Q_кон";
                lines.Cells[1, 22].Value = "Imax";
                #endregion

                #region CheckResult
                if (package.Workbook.Worksheets["Result"] == null)
                {
                    package.Workbook.Worksheets.Add("Result");
                }
                ExcelWorksheet result = package.Workbook.Worksheets["Result"];

                result.Cells[1, 1].Value = "Место приложения управляющего воздействия";
                result.Cells[2, 1, 3, 1].Merge = true;
                result.Cells[2, 1].Value = "Отключение";
                result.Cells[2, 2, 3, 2].Merge = true;
                result.Cells[2, 2].Value = "Номер начала";
                result.Cells[2, 3, 3, 3].Merge = true;
                result.Cells[2, 3].Value = "Номер конца";

                result.Cells[4, 1].Value = "Нормальный режим";
                #endregion

                DateTime dateTimeNow = DateTime.Now;

                if (!Directory.Exists(@"Results")) Directory.CreateDirectory(@"Results");
                GetPathFile = $@"Results\Solved {dateTimeNow:dd.MM.yyyy HH.mm.ss}.xlsx";
                fi = new FileInfo(GetPathFile);
                package.SaveAs(fi);
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



        public void OutputResult(Matrix EstimatedVoltage, Matrix EstimatedAngle, Matrix dP, Matrix dQ, double baseNumber, double baseVoltage, Node[] nodesArray, Line[] linesArray, int step, Matrix TanFi)
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

                //Заполнение информации о возмущении и месте воздействия
                result.Cells[1, 1, 1, nodesArray.Length + 2].Merge = true;
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

                //Вывод напряжений
                j = 0;
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
                    }
                }
                step++;

                package.Save();

                
            }

        }
    }
}
