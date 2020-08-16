using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SteadyStateCalculation
{
    class Complex
    {
        double module;
        double angle;

        public Complex() { }

        public Complex(Complex complex)
        {
            Real = complex.Real;
            Image = complex.Image;
        }

        public Complex(double real, double image)
        {
            Real = real;
            Image = image;
        }
        public double Real { get; set; }
        public double Image { get; set; }
        public double Module 
        {
            get
            {
                module = Math.Pow((Real * Real + Image * Image), 0.5);
                return module;
            }
        }
        public double Angle
        {
            get
            {
                angle = Math.Atan2(Image, Real)*180/Math.PI;
                return angle;
            }
        }
        
        public static Complex operator +(Complex num1, Complex num2)
        {
            Complex result = new Complex();

            result.Real = num1.Real + num2.Real;
            result.Image = num1.Image + num2.Image;

            return result;
        }

        public static Complex operator -(Complex num1, Complex num2)
        {
            Complex result = new Complex();

            result.Real = num1.Real - num2.Real;
            result.Image = num1.Image - num2.Image;

            return result;
        }

        public static Complex operator *(Complex num1, Complex num2)
        {
            Complex result = new Complex();

            result.Real = num1.Real*num2.Real - num1.Image*num2.Image;
            result.Image = num1.Real*num2.Image+num1.Image*num2.Real;

            return result;
        }

        public static Complex operator *(Complex num1, double factor)
        {
            Complex result = new Complex();

            result.Real = num1.Real*factor;
            result.Image = num1.Image*factor;

            return result;
        }

        public static Complex operator *(double factor, Complex num1)
        {
            Complex result = new Complex();

            result.Real = num1.Real * factor;
            result.Image = num1.Image * factor;

            return result;
        }

        public static Complex operator /(Complex num1, Complex num2)
        {
            Complex result = new Complex();

            result.Real = (num1.Real * num2.Real + num1.Image * num2.Image)/(num2.Real*num2.Real+num2.Image*num2.Image);
            result.Image = (num1.Image * num2.Real - num1.Real * num2.Image) / (num2.Real * num2.Real + num2.Image * num2.Image);

            return result;
        }

        public static Complex operator /(Complex num1, double factor)
        {
            Complex result = new Complex();

            result.Real = num1.Real/factor;
            result.Image = num1.Image/factor;

            return result;
        }

        public static bool operator ==(Complex num1, Complex num2)
        {
            if ((num1.Real == num2.Real) & (num1.Image == num2.Image)) return true;
            else return false;
        }
        public static bool operator !=(Complex num1, Complex num2)
        {
            if ((num1.Real == num2.Real) & (num1.Image == num2.Image)) return false;
            else return true;
        }

        public override bool Equals(object obj)
        {
            return obj is Complex complex &&
                   module == complex.module &&
                   angle == complex.angle &&
                   Real == complex.Real &&
                   Image == complex.Image &&
                   Module == complex.Module &&
                   Angle == complex.Angle;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(module, angle, Real, Image, Module, Angle);
        }


        public static Complex Conjugate(Complex num)
        {
            Complex result = new Complex();
            result.Real = num.Real;
            result.Image = (-1.0) * num.Image;
            return result;
        }

        public override string ToString()
        {
            if (this.Image < 0)
            {
                return $"{this.Real:0.####}-j{-this.Image:0.####}";

            }
            else
            {
                return $"{this.Real:0.####}+j{this.Image:0.####}";
            }
        }

    }
}
