using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace StudentInfoSystem.Models
{
    public class StudentDBOperation
    {
        string connectionString = @"Data Source = ICS-12; Initial Catalog = StudentTable; Integrated Security=True;";

        public List<StudentInfo> ListAll()
        {
            List<StudentInfo> lst = new List<StudentInfo>();
            using (SqlConnection sqlCon = new SqlConnection(connectionString))
            {
                sqlCon.Open();
                SqlCommand sqlCom = new SqlCommand("StudentList_SP", sqlCon);
                sqlCom.CommandType = CommandType.StoredProcedure;
                sqlCom.ExecuteNonQuery();
                SqlDataReader rdr = sqlCom.ExecuteReader();
                while (rdr.Read())
                {
                    lst.Add(new StudentInfo
                    {
                        StudentID = Convert.ToInt32(rdr["StudentID"]),
                        StudentName = rdr["StudentName"].ToString(),
                        DateOfBirth = Convert.ToDateTime(rdr["DateOfBirth"]),
                        Gender = rdr["Gender"].ToString(),
                        BloodGroup = rdr["BloodGroup"].ToString(),
                        Reliogion =rdr["Reliogion"].ToString(),
                        MaritalStatus = rdr["MaritalStatus"].ToString(),
                        
                    });
                }
                return lst;
            }
        }

        //Method for Adding an Product 
        public int AddProduct(StudentInfo student)
        {
            int i;
            using (SqlConnection sqlCon = new SqlConnection(connectionString))
            {
                sqlCon.Open();

                SqlCommand com = new SqlCommand("StudentInsert_SP", sqlCon);
                com.CommandType = CommandType.StoredProcedure;
                com.Parameters.AddWithValue("@StudentName", student.StudentName);
                com.Parameters.AddWithValue("@DateOfBirth", student.DateOfBirth.Date);
                com.Parameters.AddWithValue("@Gender", student.Gender);
                com.Parameters.AddWithValue("@BloodGroup", student.BloodGroup);
                com.Parameters.AddWithValue("@Reliogion", student.Reliogion);
                com.Parameters.AddWithValue("@MaritalStatus", student.MaritalStatus);
                com.Parameters.AddWithValue("@IsAdmin", student.IsAdmin);
                com.Parameters.AddWithValue("@Interest", student.Interest);
                com.Parameters.AddWithValue("@RegisteredAt", DateTime.Now);

                i = com.ExecuteNonQuery();
            }
            return i;
        }

    }
}