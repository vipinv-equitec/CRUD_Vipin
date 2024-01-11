using CRUDWithDapper.Models;
using Dapper;
using DocumentFormat.OpenXml.Spreadsheet;
using PagedList;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Configuration;
using System.Web.UI.WebControls;
using static Dapper.SqlMapper;

namespace CRUDWithDapper.Repository
{
    public class StudentRepository
    {
        private readonly string _connectionString;
        public StudentRepository(string connectionString)
        {
            this._connectionString = connectionString;
        }
        public void AddStudent(int studentId, string studName, int studAge, string studEmail, string studDepartment, string skills)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    db.Open();
                    var parameters = new
                    {
                        StudentId = studentId,
                        StudName = studName,
                        StudAge = studAge,
                        StudEmail = studEmail,
                        StudDepartment = studDepartment,
                        Skills = skills
                    };
                    db.Execute("StudentEdit", parameters, commandType: CommandType.StoredProcedure);
                    db.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public List<Student> GetAllStudent()
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    db.Open();
                    IList<Student> StudList = SqlMapper.Query<Student>(db, "StudentViewAll").ToList();
                    db.Close();
                    return StudList.ToList();
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public void UpdateStudent(int studentId, string studName, int studAge, string studEmail, string studDepartment, string skills)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    db.Open();
                    var parameters = new
                    {
                        StudentId = studentId,
                        StudName = studName,
                        StudAge = studAge,
                        StudEmail = studEmail,
                        StudDepartment = studDepartment,
                        Skills = skills
                    };
                    db.Execute("StudentEdit", parameters, commandType: CommandType.StoredProcedure);
                    db.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public bool DeleteStudent(int id)
        {
            try
            {
                DynamicParameters dynamicParameters = new DynamicParameters();
                dynamicParameters.Add("@StudentId", id);
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    db.Open();
                    db.Execute("StudentDeleteByyId", dynamicParameters, commandType: CommandType.StoredProcedure);
                    db.Close();
                    return true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public void StudentDetail(int id)
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    db.Open();
                    db.Execute("StudentViewById", id, commandType: CommandType.StoredProcedure);
                    db.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public void DeleteSoft(int id)
        {
            try
            {
                DynamicParameters dynamicParameters = new DynamicParameters();
                dynamicParameters.Add("@StudentId", id);
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    db.Open();
                    db.Execute("DeleteSoft", dynamicParameters, commandType: CommandType.StoredProcedure);
                    db.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public List<Student> RestoreStudent()
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    db.Open();
                    IList<Student> StudList = SqlMapper.Query<Student>(db, "RestoreStudent").ToList();
                    db.Close();
                    return StudList.ToList();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public void BackStudent(int id)
        {
            try
            {
                DynamicParameters dynamicParameters = new DynamicParameters();
                dynamicParameters.Add("@StudentId", id);
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    db.Open();
                    db.Execute("BackStudent", dynamicParameters, commandType: CommandType.StoredProcedure);
                    db.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public StringBuilder PrintData()
        {
            try
            {
                using (IDbConnection db = new SqlConnection(_connectionString))
                {
                    db.Open();
                    IList<Student> StudList = SqlMapper.Query<Student>(db, "StudentViewAll").ToList();
                    db.Close();
                    StringBuilder stringBuilder = new StringBuilder();
                    stringBuilder.AppendLine("<html><head><title>Table Print</title></head><body>");
                    stringBuilder.AppendLine("<style>table { border-collapse: collapse; width: 100%; } th, td { border: 1px solid black; padding: 8px; text-align: left; }</style>");
                    stringBuilder.AppendLine("<table class='table table-bordered'><thead><tr><th>StudentId</th><th>Name</th><th>Age</th><th>Email</th><th>Department</th><th>Skills</th></tr></thead>");

                    foreach (var student in StudList)
                    {
                        stringBuilder.AppendLine("<tr>");
                        stringBuilder.AppendLine($"<td>{student.StudentId}</td>");
                        stringBuilder.AppendLine($"<td>{student.StudName}</td>");
                        stringBuilder.AppendLine($"<td>{student.StudAge}</td>");
                        stringBuilder.AppendLine($"<td>{student.StudEmail}</td>");
                        stringBuilder.AppendLine($"<td>{student.StudDepartment}</td>");
                        stringBuilder.AppendLine($"<td>{student.Skills}</td>");
                    }
                    stringBuilder.AppendLine("</tbody></table></body></html>");
                    return stringBuilder;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public List<Student> GetStudents(int page, int pageSize)
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                db.Open();
                IList<Student> StudList = SqlMapper.Query<Student>(db, "StudentViewAll").ToList();
                var offset = (page - 1) * pageSize;
                var query = $"SELECT * FROM Student where IsActive = 1 ORDER BY StudName OFFSET {offset} ROWS FETCH NEXT {pageSize} ROWS ONLY";

                db.Close();

                return (List<Student>)db.Query<Student>(query);
            }
        }
        public int GetTotalStudentCount()
        {
            using (IDbConnection db = new SqlConnection(_connectionString))
            {
                db.Open();
                var countQuery = "SELECT COUNT(*) FROM Student where IsActive = 1";

                db.Close();
                return db.ExecuteScalar<int>(countQuery);
            }
        }
    }
}