using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StudentEvaluationSystem
{
    internal class Student
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Department { get; set; }
        public int TheoryScore { get; set; }
        public int PracticalScore { get; set; }
        public string Phone {  get; set; }
        public string Email { get; set; }
        public double FinalScore { get; set; }
        public string Message { get; set; }
    }
}
