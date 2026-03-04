using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
namespace Pinnacle.Models
{
    internal class Validate
    {
        
        public bool IsInteger(String strNumber)
        {
            Regex objNotIntPattern = new Regex("[^0-9-]");
            Regex objIntPattern = new Regex("^-[0-9-]+$|^[0-9-]+$");
            return !objNotIntPattern.IsMatch(strNumber) &&
            objIntPattern.IsMatch(strNumber);
        }
        public bool IsIntegerdot(String strNumber)
        {
            Regex objNotIntPattern = new Regex("[^0-9.]");
            Regex objIntPattern = new Regex("^-[0-9.]+$|^[0-9.]+$");
            return !objNotIntPattern.IsMatch(strNumber) &&
            objIntPattern.IsMatch(strNumber);
        }
        public bool IsIntegerSpacehypen(String strNumber)
        {
            Regex objNotIntPattern = new Regex("[^ 0-9-]");
            Regex objIntPattern = new Regex("^-[ 0-9-]+$|^[ 0-9- ]+$");
            return !objNotIntPattern.IsMatch(strNumber) &&
            objIntPattern.IsMatch(strNumber);
        }

        public bool IsDecimal(string strNumber)
        {
            Regex objNotIntPattern = new Regex("[^.0-9-.]");
            Regex objIntPattern = new Regex("^-[.0-9.]+$|^[.0-9.]+$");
            return !objNotIntPattern.IsMatch(strNumber) &&
            objIntPattern.IsMatch(strNumber);
        }

        public bool IsStringSpace(String strNumber)
        {
            Regex objNotIntPattern = new Regex("[^a-zA-Z ]");
            Regex objIntPattern = new Regex("^-[a-zA-Z ]+$|^[a-zA-Z ]+$");
            return !objNotIntPattern.IsMatch(strNumber) &&
            objIntPattern.IsMatch(strNumber);
        }
        public bool IsStringSpacepracket(String strNumber)
        {
            Regex objNotIntPattern = new Regex("[^a-zA-Z ()]");
            Regex objIntPattern = new Regex("^-[a-zA-Z ()]+$|^[a-zA-Z ()]+$");
            return !objNotIntPattern.IsMatch(strNumber) &&
            objIntPattern.IsMatch(strNumber);
        }
        public bool IsStringSpacehypenpracket(String strNumber)
        {
            Regex objNotIntPattern = new Regex("[^a-zA-Z ()-]");
            Regex objIntPattern = new Regex("^-[a-zA-Z ()-]+$|^[a-zA-Z ()-]+$");
            return !objNotIntPattern.IsMatch(strNumber) &&
            objIntPattern.IsMatch(strNumber);
        }
        public bool IsStringSpacehypen(String strNumber)
        {
            Regex objNotIntPattern = new Regex("[^a-zA-Z -]");
            Regex objIntPattern = new Regex("^-[a-zA-Z -]+$|^[a-zA-Z -]+$");
            return !objNotIntPattern.IsMatch(strNumber) &&
            objIntPattern.IsMatch(strNumber);
        }
        public bool IsString(String strNumber)
        {
            Regex objNotIntPattern = new Regex("[^a-zA-Z]");
            Regex objIntPattern = new Regex("^-[a-zA-Z]+$|^[a-zA-Z]+$");
            return !objNotIntPattern.IsMatch(strNumber) &&
            objIntPattern.IsMatch(strNumber);
        }
        public bool IsStringNumberic(String strNumber)
        {
            Regex objNotIntPattern = new Regex("[^0-9a-zA-Z]");
            Regex objIntPattern = new Regex("^-[0-9a-zA-Z]+$|^[0-9a-zA-Z]+$");
            return !objNotIntPattern.IsMatch(strNumber) &&
            objIntPattern.IsMatch(strNumber);
        }
        public bool IsStringNumbericlace(String strNumber)
        {
            Regex objNotIntPattern = new Regex("[^0-9a-zA-Z/]");
            Regex objIntPattern = new Regex("^-[0-9a-zA-Z/]+$|^[0-9a-zA-Z/]+$");
            return !objNotIntPattern.IsMatch(strNumber) &&
            objIntPattern.IsMatch(strNumber);
        }
        public bool IsStringNumberichypen(String strNumber)
        {
            Regex objNotIntPattern = new Regex("[^0-9a-zA-Z-]");
            Regex objIntPattern = new Regex("^-[0-9a-zA-Z-]+$|^[0-9a-zA-Z-]+$");
            return !objNotIntPattern.IsMatch(strNumber) &&
            objIntPattern.IsMatch(strNumber);
        }
        public bool IsStringNumbericspace(String strNumber)
        {
            Regex objNotIntPattern = new Regex("[^0-9 a-zA-Z]");
            Regex objIntPattern = new Regex("^-[ 0-9 a-zA-Z ]+$|^[ 0-9 a-zA-Z ]+$");
            return !objNotIntPattern.IsMatch(strNumber) &&
            objIntPattern.IsMatch(strNumber);
        }
        public bool IsStringNumbericspacehypen(String strNumber)
        {
            Regex objNotIntPattern = new Regex("[^0-9a-zA-Z -]");
            Regex objIntPattern = new Regex("^-[0-9 a-zA-Z -]+$|^[0-9 a-zA-Z -]+$");
            return !objNotIntPattern.IsMatch(strNumber) &&
            objIntPattern.IsMatch(strNumber);
        }
        public bool IsStringNumbericspaceslace(String strNumber)
        {
            Regex objNotIntPattern = new Regex("[^0-9a-zA-Z /]");
            Regex objIntPattern = new Regex("^-[0-9 a-zA-Z /]+$|^[0-9 a-zA-Z /]+$");
            return !objNotIntPattern.IsMatch(strNumber) &&
            objIntPattern.IsMatch(strNumber);
        }
        public bool IsStringNumberic1(String strNumber)
        {
            Regex objNotIntPattern = new Regex("[^0-9a-zA-Z.]");
            Regex objIntPattern = new Regex("^-[0-9a-zA-Z.]+$|^[0-9a-zA-Z.]+$");
            return !objNotIntPattern.IsMatch(strNumber) &&
            objIntPattern.IsMatch(strNumber);
        }
        public bool IsStringNumbericslace(String strNumber)
        {
            Regex objNotIntPattern = new Regex("[^0-9a-zA-Z/]");
            Regex objIntPattern = new Regex("^-[0-9a-zA-Z/]+$|^[0-9a-zA-Z/]+$");
            return !objNotIntPattern.IsMatch(strNumber) &&
            objIntPattern.IsMatch(strNumber);
        }

        public bool IsStringNumbericslacehyper(String strNumber)
        {
            Regex objNotIntPattern = new Regex("[^0-9a-zA-Z-]");
            Regex objIntPattern = new Regex("^-[0-9a-zA-Z-]+$|^[0-9a-zA-Z-]+$");
            return !objNotIntPattern.IsMatch(strNumber) &&
            objIntPattern.IsMatch(strNumber);
        }
        public bool IsStringdotspace(String strNumber)
        {
            Regex objNotIntPattern = new Regex("[^a-zA-Z #]");
            Regex objIntPattern = new Regex("^-[a-zA-Z #]+$|^[a-zA-Z #]+$");
            return !objNotIntPattern.IsMatch(strNumber) &&
            objIntPattern.IsMatch(strNumber);
        }
    }
}