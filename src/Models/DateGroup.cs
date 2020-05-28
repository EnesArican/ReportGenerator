using System;

namespace ReportGenerator.Models
{
    public class DateGroup : IEquatable<DateGroup>
    {
        public int Month { get; set; }

        public int Year { get; set; }


        //Need these methods so that the object can be compared corretly inside linq statements
        public override bool Equals(object obj)
        {
            return Equals(obj as DateGroup);
        }

        public bool Equals(DateGroup other)
        {
            return other != null &&
                   Month == other.Month &&
                   Year == other.Year;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(Month, Year);
        }
    }
}
