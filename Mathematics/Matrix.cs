using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text;

namespace SteadyStateCalculation
{
    class Matrix
    {
        public double[,] matrix;

        public Matrix(int row = 1, int col = 1)
        {
            matrix = new double[row, col];
        }

        public Matrix(double[,] mat)
        {
            matrix = mat;
        }
        
        // Получение количества строк и столбцов матрицы
        public int GetRows
        {
            get
            {
                return matrix.GetLength(0);
            }
        }
        public int GetColumns 
        {
            get
            {
                return matrix.GetLength(1);
            }
        }
        
        //Сложение матриц
        public static Matrix operator +(Matrix mat1, Matrix mat2)
        {
            if ((mat1.GetRows != mat2.GetRows) | (mat1.GetColumns != mat2.GetColumns))
            {
                throw new ArgumentException("Matrices can not be summed up");
            }
            else
            {
                Matrix result = new Matrix(mat1.GetRows,mat1.GetColumns);

                for (int i = 0; i < mat1.GetRows; i++)
                {
                    for(int j = 0; j<mat1.GetColumns;j++)
                    {
                        result.matrix[i, j] = mat1.matrix[i, j] + mat2.matrix[i, j];
                    }
                }
                return result;
            }
        }
        
        //Вычитание матриц
        public static Matrix operator -(Matrix mat1, Matrix mat2)
        {
            if ((mat1.GetRows != mat2.GetRows) | (mat1.GetColumns != mat2.GetColumns))
            {
                throw new ArgumentException("Matrices can not be deducted");
            }
            else
            {
                Matrix result = new Matrix(mat1.GetRows, mat1.GetColumns);

                for (int i = 0; i < mat1.GetRows; i++)
                {
                    for (int j = 0; j < mat1.GetColumns; j++)
                    {
                        result.matrix[i, j] = mat1.matrix[i, j] - mat2.matrix[i, j];
                    }
                }
                return result;
            }
        }
        
        //Перемножение матриц
        public static Matrix operator *(Matrix mat1, Matrix mat2)
        {
            if (mat1.GetColumns != mat2.GetRows)
            {
                throw new ArgumentException("Matrices can not be multiplied");
            }
            else
            {
                Matrix result = new Matrix(mat1.GetRows, mat2.GetColumns);

                for (int i = 0; i < mat1.GetRows; i++)
                {
                    for (int j = 0; j < mat2.GetColumns; j++)
                    {
                        for(int k = 0; k<mat1.GetColumns;k++)
                        {
                            result.matrix[i, j] += mat1.matrix[i, k] * mat2.matrix[k, j];
                        }
                    }
                }
                return result;
            }
        }
        
        //Умножение элементов матрицы на число
        public static Matrix operator *(Matrix mat, double factor)
        {
            Matrix result = new Matrix(mat.GetRows, mat.GetColumns);

            for (int i = 0; i < mat.GetRows; i++)
            {
                for (int j = 0; j < mat.GetColumns; j++)
                {
                    result.matrix[i, j] = mat.matrix[i, j] * factor;
                }
            }
            return result;
        }

        public static Matrix operator *(double factor, Matrix mat)
        {
            Matrix result = mat*factor;
            return result;
        }
        
        //Правостороннее деление матриц
        public static Matrix operator /(Matrix mat1, Matrix mat2)
        {
            if (mat1.GetRows != mat2.GetColumns)
            {
                throw new ArgumentException("Matrices can not be divided");
            }
            else
            {
                return CreateInvertedMatrix(mat1)*mat2;
            }
        }
        
        //Деление элементов матрицы на число
        public static Matrix operator /(Matrix mat, double factor)
        {
            Matrix result = new Matrix(mat.GetRows, mat.GetColumns);

            for (int i = 0; i < mat.GetRows; i++)
            {
                for (int j = 0; j < mat.GetColumns; j++)
                {
                    result.matrix[i, j] = mat.matrix[i, j] / factor;
                }
            }
            return result;
        }
        
        //Нахождение определителя матрицы
        public static double FindDeterminant(Matrix mat)
        {
            if (mat.GetRows != mat.GetColumns)
            {
                throw new ArgumentException("Determinant cannot be found");
            }
            else
            {
                if(mat.GetRows==2)
                {
                    return mat.matrix[0, 0] * mat.matrix[1, 1] - mat.matrix[0, 1] * mat.matrix[1, 0];
                }
                double determinant = 0;
                for(int j=0;j<mat.GetColumns;j++)
                {
                    determinant += Math.Pow(-1, 1 + j) * mat.matrix[1, j] * FindDeterminant(CreateMinor(mat, 1, j));
                }
                return determinant;
            }
        }
        
        //Нахождение минора матрицы (исключение строки row и столбца col)
        public static Matrix CreateMinor(Matrix mat, int row, int col)
        {
            if ((mat.GetRows < 2) | (mat.GetColumns < 2))
            {
                throw new ArgumentException("Minor can not be created");
            }
            else
            {
                Matrix Minor = new Matrix(mat.GetRows - 1, mat.GetColumns - 1);
                
                int r = 0;
                int c = 0;
                for (int i = 0; i < mat.GetRows; i++)
                {
                    if (i == row)
                    {
                        continue;
                    }
                    else
                    {
                        for (int j = 0; j < mat.GetColumns; j++)
                        {
                            if (j == col)
                            {
                                continue;
                            }
                            else
                            {
                                Minor.matrix[r, c] = mat.matrix[i, j];
                            }
                            c++;
                        }
                    }
                    r++;
                    c = 0;
                }
                return Minor;
            }
        }
        
        //Создание транспонированной матрицы
        public static Matrix CreateTransposedMatrix(Matrix mat)
        {
            Matrix TransposedMatrix = new Matrix(mat.GetRows, mat.GetColumns);

            for(int i=0;i<mat.GetRows;i++)
                for(int j=0;j<mat.GetColumns;j++)
                {
                    TransposedMatrix.matrix[i, j] = mat.matrix[j, i];
                }
            return TransposedMatrix;
        }
        
        //Создание обратной матрицы
        public static Matrix CreateInvertedMatrix(Matrix mat)
        {
            if (mat.GetRows != mat.GetColumns)
            {
                throw new ArgumentException("Matrix cannot be inverted");
            }
            else
            {
                Matrix InvertedMatrix = new Matrix(mat.GetRows, mat.GetColumns);

                double determinant = FindDeterminant(mat);

                for(int i=0;i<mat.GetRows;i++)
                    for(int j=0;j<mat.GetColumns;j++)
                    {
                        InvertedMatrix.matrix[i, j] = Math.Pow(-1, i + j) * FindDeterminant(CreateMinor(mat, i, j)) / determinant;
                    }
                return CreateTransposedMatrix(InvertedMatrix);
            }

        }

        public static Matrix CreateUnitMatrix(int dimension)
        {
            if (dimension < 1)
            {
                throw new ArgumentException("UnitMatrix cannot be created. Check dimension.");
            }
            else
            {
                Matrix UnitMatrix = new Matrix(dimension, dimension);
                for (int i = 0; i < dimension; i++)
                    for (int j = 0; j < dimension; j++)
                    {
                        if (i == j) UnitMatrix.matrix[i, j] = 1.0;
                        else UnitMatrix.matrix[i, j] = 0.0;
                    }
                return UnitMatrix;
            }
        }

        public static Matrix CreateDiagMatrix(Matrix VectorMatrix)
        {
            if (VectorMatrix.GetRows > VectorMatrix.GetColumns)
            {
                Matrix DiagMatrix = new Matrix(VectorMatrix.GetRows, VectorMatrix.GetRows);
                for (int i = 0; i < DiagMatrix.GetRows; i++)
                    for (int j = 0; j < DiagMatrix.GetColumns; j++)
                    {
                        if (i == j) DiagMatrix.matrix[i, j] = VectorMatrix.matrix[i,0];
                        else DiagMatrix.matrix[i, j] = 0.0;
                    }
                return DiagMatrix;
            }
            else
            {
                Matrix DiagMatrix = new Matrix(VectorMatrix.GetColumns, VectorMatrix.GetColumns);
                for (int i = 0; i < DiagMatrix.GetRows; i++)
                    for (int j = 0; j < DiagMatrix.GetColumns; j++)
                    {
                        if (i == j) DiagMatrix.matrix[i, j] = VectorMatrix.matrix[0, i];
                        else DiagMatrix.matrix[i, j] = 0.0;
                    }
                return DiagMatrix;
            }
        }

        public static Matrix CopyMatrix(Matrix MatrixForCopy)
        {
            Matrix NewMatrix = new Matrix(MatrixForCopy.GetRows, MatrixForCopy.GetColumns);
            for (int i = 0; i < NewMatrix.GetRows; i++)
            {
                for (int j = 0; j < NewMatrix.GetColumns; j++)
                {
                    NewMatrix.matrix[i, j] = MatrixForCopy.matrix[i, j];
                }
            }
            return NewMatrix;
        }

        public static Matrix CreateInvertedMatrixGauss(Matrix mat)
        {
            if (mat.GetRows != mat.GetColumns)
            {
                throw new ArgumentException("MatrixComplex cannot be inverted");
            }
            else
            {
                Matrix GaussMatrix = new Matrix(mat.GetRows, mat.GetColumns);
                GaussMatrix = CopyMatrix(mat);
                Matrix JointMatrix = CreateUnitMatrix(mat.GetRows);
                Matrix TranspositionMatrix = CreateUnitMatrix(mat.GetRows);

                //Прямой ход
                for(int j = 0; j<GaussMatrix.GetColumns;j++)
                {
                    for(int i=j;i<GaussMatrix.GetRows;i++)
                    {
                        if(GaussMatrix.matrix[j,j]==0)
                        {
                            int newRow = FindNotZero(GaussMatrix, j, j, false);
                            SwapRows(GaussMatrix, i, newRow);
                            SwapRows(JointMatrix, i, newRow);
                            SwapRows(TranspositionMatrix, i, newRow);
                        }
                        if (i == j) continue;
                        SummRows(JointMatrix, j, i, (-1.0) * GaussMatrix.matrix[i, j] / GaussMatrix.matrix[j, j]);
                        SummRows(GaussMatrix, j, i, (-1.0) * GaussMatrix.matrix[i, j] / GaussMatrix.matrix[j, j]);
                    }
                }
                //Обратный ход
                for (int j = GaussMatrix.GetColumns - 1; j > -1; j--)
                {
                    for (int i = j; i > -1; i--)
                    {
                        if (GaussMatrix.matrix[j, j] == 0)
                        {
                            int newRow = FindNotZero(GaussMatrix, j, j, true);
                            SwapRows(GaussMatrix, i, newRow);
                            SwapRows(JointMatrix, i, newRow);
                            SwapRows(TranspositionMatrix, i, newRow);
                        }
                        if (i == j) continue;
                        SummRows(JointMatrix, j, i, (-1.0) * GaussMatrix.matrix[i, j] / GaussMatrix.matrix[j, j]);
                        SummRows(GaussMatrix, j, i, (-1.0) * GaussMatrix.matrix[i, j] / GaussMatrix.matrix[j, j]);
                    }
                }
                for (int i = 0; i < GaussMatrix.GetRows; i++)
                {
                    MultiplyRow(JointMatrix, i, GaussMatrix.matrix[i, i]);
                    MultiplyRow(GaussMatrix, i, GaussMatrix.matrix[i, i]);
                }
                return (JointMatrix); //Почему не на матрицу перестановок?

                static int FindNotZero(Matrix GaussMatrix, int currentRow, int currentColumn, bool direction)
                {
                    //direction = false(down); direction = true(up)
                    bool IsFound = false;
                    int row = currentRow;
                    while (!IsFound)
                    {
                        if (GaussMatrix.matrix[row, currentColumn] == 0)
                        {
                            if (direction)
                            {
                                if (row == 0)
                                {
                                    Console.WriteLine("Обнаружена сингулярность");
                                    break;
                                }
                                else row--;
                            }
                            else
                            {
                                if (row == GaussMatrix.GetRows - 1)
                                {
                                    Console.WriteLine("Обнаружена сингулярность");
                                    break;
                                }
                                else row++;
                            }
                        }
                        else IsFound = true;
                    }
                    return row;
                }

            }
        }

        public static void SwapRows(Matrix mat, int row1, int row2)
        {
            double temp;
            for(int j=0;j<mat.GetColumns;j++)
            {
                temp = mat.matrix[row1, j];
                mat.matrix[row1, j] = mat.matrix[row2, j];
                mat.matrix[row2, j] = temp;
            }
        }

        public static Matrix SummRows(Matrix mat, int baseRow, int sumRow, double factor)
        {
            for (int j = 0; j < mat.GetColumns; j++)
            {
                mat.matrix[sumRow, j] += mat.matrix[baseRow, j] * factor;
            }
            return mat;
        }

        public static Matrix MultiplyRow(Matrix mat, int baseRow, double factor)
        {
            for (int j = 0; j < mat.GetColumns; j++)
            {
                mat.matrix[baseRow, j] = mat.matrix[baseRow, j] / factor;
            }
            return mat;
        }

        public static Matrix ConjugateMatrixElements(Matrix mat)
        {
            Matrix ConjugateM = new Matrix(mat.GetRows, mat.GetColumns);
            ConjugateM = CopyMatrix(mat);

            for (int i = ConjugateM.GetRows/2; i < ConjugateM.GetRows; i++)
                ConjugateM.matrix[i, 0] = (-1.0)*ConjugateM.matrix[i, 0];
            return ConjugateM;
        }

        //Вывод матрицы в виде строки
        public override string ToString()
        {
            StringBuilder ret = new StringBuilder();

            if (matrix == null)
            {
                ret.Append("Пустая матрица или матрица отсутствует");
                return ret.ToString();
            }
            else
            {
                for(int i=0;i<GetRows;i++)
                {
                    for(int j=0;j<GetColumns;j++)
                    {
                        ret.Append($"{matrix[i, j],-8:0.####}");
                    }
                    ret.Append("\n");
                }
            }
            return ret.ToString();
        }
    }
}
