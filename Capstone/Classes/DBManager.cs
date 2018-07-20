using Capstone.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Linq;
using System.Web;

namespace Capstone.Classes
{
    public class DBManager
    {
        public OleDbConnection GetConnection()
        {
            OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|/CapstoneDatabase.accdb");
            return connection;
        }
        public List<History> GetHistoryList()
        {
            List<History> historyList = new List<History>();
            OleDbConnection connection = GetConnection();
            connection.Open();
            OleDbDataReader reader = null;
            OleDbCommand command = new OleDbCommand("SELECT * FROM  History ORDER BY [GenDate] DESC", connection);
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                History h = new History();

                h.FileName = reader["FileName"].ToString();
                h.FileURL = reader["FileURL"].ToString();
                h.ExamFileName = reader["ExamFileName"].ToString();
                h.ExamURL = reader["ExamURL"].ToString();
                h.CalendarURL = reader["CalendarURL"].ToString();
                h.GenDate = DateTime.Parse(reader["GenDate"].ToString());
                h.User = reader["User"].ToString();

                historyList.Add(h);
            }
            connection.Close();
            return historyList;
        }
        public void NewHistoryEntry(string file, string fileURL, string exam, string examURL, string calendarURL, string user)
        {
            OleDbConnection connection = GetConnection();
            OleDbCommand command = new OleDbCommand("INSERT INTO History([FileName], [FileURL], [ExamFileName], [ExamURL], [CalendarURL], [GenDate], [User]) VALUES (@file,@fileURL,@exam,@examURL,@calURL,@genDate,@user)", connection);
            command.Parameters.AddWithValue("@file", file);
            command.Parameters.AddWithValue("@fileURL", fileURL);
            command.Parameters.AddWithValue("@exam", exam);
            command.Parameters.AddWithValue("@examURL", examURL);
            command.Parameters.AddWithValue("@calURL", calendarURL);
            command.Parameters.AddWithValue("@genDate", DateTime.Now.ToString());
            command.Parameters.AddWithValue("@user", user);
            try
            {
                connection.Open();
                command.ExecuteNonQuery();
            }
            catch(Exception e){ }
            connection.Close();
        }
        public void SavePreferences(string user, string fileDir, string examDir, string calGen)
        {

            // Fix directory strings
            if (fileDir[0].ToString() != "/")
                fileDir = "/" + fileDir;
            if (fileDir[fileDir.Length-1].ToString() != "/")
                fileDir = fileDir + "/";

            if (examDir[0].ToString() != "/")
                examDir = "/" + examDir;
            if (examDir[examDir.Length - 1].ToString() != "/")
                examDir = examDir + "/";

            OleDbConnection connection = GetConnection();
            try
            {
                connection.Open();

                // ------------------------------------
                // Upload Directory Section
                OleDbCommand command = new OleDbCommand("SELECT * FROM  UploadDirectory WHERE [User] LIKE @user", connection);
                command.Parameters.AddWithValue("@user", user);
                OleDbDataReader reader = null;
                reader = command.ExecuteReader();
                if (reader.Read())
                {
                    command = new OleDbCommand("UPDATE UploadDirectory SET [Directory] = @dir WHERE [User] = @user", connection);
                    command.Parameters.AddWithValue("@dir", fileDir);
                    command.Parameters.AddWithValue("@user", user);

                    command.ExecuteNonQuery();
                }
                else
                {
                    command = new OleDbCommand("INSERT INTO UploadDirectory ([User], [Directory]) VALUES (@user,@dir)", connection);
                    command.Parameters.AddWithValue("@user", user);
                    command.Parameters.AddWithValue("@dir", fileDir);

                    command.ExecuteNonQuery();
                }
                // ------------------------------------
                // Exam Directory Section
                command = new OleDbCommand("SELECT * FROM  ExamDirectory WHERE [User] = @user", connection);
                command.Parameters.AddWithValue("@user", user);
                reader = null;
                reader = command.ExecuteReader();
                if (reader.Read())
                {
                    command = new OleDbCommand("UPDATE ExamDirectory SET [Directory] = @dir WHERE [User] LIKE @user", connection);
                    command.Parameters.AddWithValue("@dir", examDir);
                    command.Parameters.AddWithValue("@user", user);

                    command.ExecuteNonQuery();
                }
                else
                {
                    command = new OleDbCommand("INSERT INTO ExamDirectory ([User], [Directory]) VALUES (@user,@dir)", connection);
                    command.Parameters.AddWithValue("@user", user);
                    command.Parameters.AddWithValue("@dir", examDir);

                    command.ExecuteNonQuery();
                }
                // ------------------------------------
                // Calendar Generation Section
                command = new OleDbCommand("SELECT * FROM GenerateCalendars WHERE [User] = @user", connection);
                command.Parameters.AddWithValue("@user", user);
                reader = null;
                reader = command.ExecuteReader();
                if (reader.Read())
                {
                    command = new OleDbCommand("UPDATE GenerateCalendars SET [Flag] = @flag WHERE [User] LIKE @user", connection);
                    command.Parameters.AddWithValue("@flag", calGen);
                    command.Parameters.AddWithValue("@user", user);

                    command.ExecuteNonQuery();
                }
                else
                {
                    command = new OleDbCommand("INSERT INTO GenerateCalendars ([User], [Flag]) VALUES (@user,@flag)", connection);
                    command.Parameters.AddWithValue("@user", user);
                    command.Parameters.AddWithValue("@flag", calGen);

                    command.ExecuteNonQuery();
                }
                // ------------------------------------
            }
            catch (Exception e) { }
            connection.Close();
        }
        public string GetUserUploadDirectory(string user)
        {
            OleDbConnection connection = GetConnection();
            connection.Open();
            OleDbCommand command = new OleDbCommand("SELECT * FROM  UploadDirectory WHERE [User] LIKE @user", connection);
            command.Parameters.AddWithValue("@user", user);

            OleDbDataReader reader = null;
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                string data = reader["Directory"].ToString();
                connection.Close();
                return data;
            }
            return "/Course Sheets/";
        }
        public string GetUserExamDirectory(string user)
        {
            OleDbConnection connection = GetConnection();
            connection.Open();
            OleDbCommand command = new OleDbCommand("SELECT * FROM  ExamDirectory WHERE [User] LIKE @user", connection);
            command.Parameters.AddWithValue("@user", user);

            OleDbDataReader reader = null;
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                string data = reader["Directory"].ToString();
                connection.Close();
                return data;
            }
            return "/Exam Sheets/";
        }
        public string GetUserGenCalendars(string user)
        {
            OleDbConnection connection = GetConnection();
            connection.Open();
            OleDbCommand command = new OleDbCommand("SELECT * FROM  GenerateCalendars WHERE [User] LIKE @user", connection);
            command.Parameters.AddWithValue("@user", user);

            OleDbDataReader reader = null;
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                string data = reader["Flag"].ToString();
                connection.Close();
                return data;
            }
            return "False";
        }
    }
}