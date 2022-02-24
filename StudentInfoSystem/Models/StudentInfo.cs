using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace StudentInfoSystem.Models
{
    public class StudentInfo
    {
        public int StudentID { get; set; }

        public string StudentName { get; set; }

        public DateTime DateOfBirth { get; set; }

        public string Gender { get; set; }

        public string BloodGroup { get; set; }

        public string Reliogion { get; set; }

        public string MaritalStatus { get; set; }

        public bool IsAdmin { get; set; }

        public string Interest { get; set; }

        public DateTime RegisteredAt { get; set; }


        public String DateOfBirthSt
        {
            get
            {
                if (this.DateOfBirth == DateTime.MinValue)
                {
                    return "";
                }
                else
                {
                    return this.DateOfBirth.ToString("dd MMM yyyy");
                }
            }
        }

    }
}