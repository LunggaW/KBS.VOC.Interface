using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using NLog;
using NLog.LayoutRenderers.Wrappers;

namespace KBS.RANCH.VOCOLLECT.INTERFACE.MODEL
{
    public class IdleEmail
    {
        private String ConnectionStringSqlServer = ConfigurationManager.AppSettings["ConnectionStringSqlServer"];

        private Int32 IntervalDay;
        private Int32 IntervalHour;
        private Int32 IntervalMinute;
        private Int32 IntervalSecond;
        private QuartzParam quartzParam;

        private static Logger logger = LogManager.GetCurrentClassLogger();
        private String ErrorString;

        //private OracleConnection con;
        private SqlConnection conSql;

        public IdleEmail()
        {

            try
            {
                IntervalDay = GetIntervalDay();
                IntervalHour = GetIntervalHour();
                IntervalMinute = GetIntervalMinute();
                IntervalSecond = GetIntervalSecond();
            }

            catch (Exception ex)
            {
                logger.Error("Constructor");
                logger.Error(ex.Message);
                throw;
            }
        }


        public QuartzParam getQuartzParam()
        {
            try
            {
                this.ConnectSQLServer();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = conSql;
                cmd.CommandText = "SELECT [SCHID] " +
                                  ",[SCHSDESC] " +
                                  ",[SCHLDESC] " +
                                  ",[SCHINTSEC] " +
                                  ",[SCHINTMIN] " +
                                  ",[SCHINTHOR] " +
                                  ",[SCHINTDAY] " +
                                  ",[SCHSTRTHOR] " +
                                  ",[SCHSTRTMIN] " +
                                  ",[SCHISDLY] " +
                                  ",[SCHISMTH] " +
                                  ",[SCHISWKY] " +
                                  ",[SCHDAYWK] " +
                                  ",[SCHDAYMTH] " +
                                  "FROM [MSTSCH] " +
                                  "where SCHSDESC = 'Idle Email'";

                cmd.CommandType = CommandType.Text;

                logger.Debug(cmd.CommandText);

                SqlDataReader dr = cmd.ExecuteReader();

                quartzParam = new QuartzParam();

                while (dr.Read())
                {
                    logger.Debug(dr["SCHID"].ToString() + "   " + dr["SCHSDESC"].ToString() + "   " + dr["SCHLDESC"].ToString() + "   " +
                                 dr["SCHINTSEC"].ToString() + "   " + dr["SCHINTMIN"].ToString() + "   " + dr["SCHINTHOR"].ToString() + "   " +
                                 dr["SCHINTDAY"].ToString() + "   " + dr["SCHSTRTHOR"].ToString() + "   " + dr["SCHSTRTMIN"].ToString() + "   " +
                                 dr["SCHISDLY"].ToString() + "   " + dr["SCHISMTH"].ToString() + "   " + dr["SCHISWKY"].ToString() + "   " +
                                 dr["SCHDAYWK"].ToString() + "   " + dr["SCHDAYMTH"].ToString());
                    quartzParam.ID = int.Parse(dr["SCHID"].ToString());
                    logger.Debug("Masuk ID");
                    quartzParam.ShortDesc = dr["SCHSDESC"].ToString();
                    logger.Debug("Masuk Short Desc");
                    quartzParam.LongDesc = dr["SCHLDESC"].ToString();
                    logger.Debug("Masuk Long Desc");
                    quartzParam.IntervalSecond = int.Parse(dr["SCHINTSEC"].ToString());
                    logger.Debug("Masuk Interval Secong");
                    quartzParam.IntervalMinute = int.Parse(dr["SCHINTMIN"].ToString());
                    logger.Debug("Masuk Interval Minute");
                    quartzParam.IntervalHour = int.Parse(dr["SCHINTHOR"].ToString());
                    logger.Debug("Masuk Intervcal Hour");
                    quartzParam.IntervalDay = int.Parse(dr["SCHINTDAY"].ToString());
                    logger.Debug("Masuk Interval Day");
                    quartzParam.StartHour = int.Parse(dr["SCHSTRTHOR"].ToString());
                    logger.Debug("Masuk Start Hour");
                    quartzParam.StartMinute = int.Parse(dr["SCHSTRTMIN"].ToString());
                    logger.Debug("Masuk Start Min");
                    quartzParam.IsDaily = bool.Parse(dr["SCHISDLY"].ToString());
                    logger.Debug("Masuk Is Daily");
                    quartzParam.IsMonthly = bool.Parse(dr["SCHISMTH"].ToString());
                    logger.Debug("Masuk Is Monthly");
                    quartzParam.IsWeekly = bool.Parse(dr["SCHISWKY"].ToString());
                    logger.Debug("Masuk Is Weekly");
                    quartzParam.DayOfWeek = dr["SCHDAYWK"].ToString();
                    logger.Debug("Masuk Day of Week");
                    quartzParam.DayOfMonth = int.Parse(dr["SCHDAYMTH"].ToString());
                    logger.Debug("Masuk Day of Month");

                }

                return quartzParam;

            }
            catch (Exception ex)
            {
                logger.Error("getQuartzParam function");
                logger.Error(ex.Message);
                logger.Error("Error getting GetIntervalDay Value, returning 0 as default");
                return null;
            }
        }


        public void ConnectSQLServer()
        {
            try
            {
                logger.Trace("Start Starting Connection Server");
                conSql = new SqlConnection();
                conSql.ConnectionString = ConnectionStringSqlServer;
                logger.Debug("Connection String : " + conSql.ConnectionString.ToString());
                conSql.Open();
                logger.Debug("End Starting Connection Server");
            }
            catch (SqlException ex)
            {
                logger.Error("ConnectSQLServer Function, Sql Exception");
                logger.Error(ex.Message);
                throw;
            }
            catch (Exception ex)
            {
                logger.Error("ConnectSQLServer Function");
                logger.Error(ex.Message);
                throw;
            }
        }

        public Int32 GetIntervalDay()
        {
            try
            {
                return Int32.Parse(ConfigurationManager.AppSettings["IntervalDay"]); ;
            }
            catch (Exception ex)
            {
                logger.Error("GetIntervalDay function");
                logger.Error(ex.Message);
                logger.Error("Error getting GetIntervalDay Value, returning 0 as default");
                return 0;
            }
        }

        public Int32 GetIntervalHour()
        {
            try
            {
                return Int32.Parse(ConfigurationManager.AppSettings["IntervalHour"]); ;
            }
            catch (Exception ex)
            {
                logger.Error("GetIntervalHour function");
                logger.Error(ex.Message);
                logger.Error("Error getting GetIntervalHour Value, returning 0 as default");
                return 0;
            }
        }

        public Int32 GetIntervalMinute()
        {
            try
            {
                return Int32.Parse(ConfigurationManager.AppSettings["IntervalMinute"]); ;
            }
            catch (Exception ex)
            {
                logger.Error("GetIntervalMinute function");
                logger.Error(ex.Message);
                logger.Error("Error getting GetIntervalMinute Value, returning 0 as default");
                return 0;
            }
        }

        public Int32 GetIntervalSecond()
        {
            try
            {
                return Int32.Parse(ConfigurationManager.AppSettings["IntervalSecond"]); ;
            }
            catch (Exception ex)
            {
                logger.Error("GetIntervalSecond function");
                logger.Error(ex.Message);
                logger.Error("Error getting GetIntervalSecond Value, returning 0 as default");
                return 0;
            }
        }

        public void CloseSqlServer()
        {
            try
            {
                logger.Debug("Closing Connection");
                conSql.Close();
                conSql.Dispose();
                logger.Debug("End Close Connection");
            }
            catch (Exception e)
            {
                logger.Error("Close Function");
                logger.Error(e.Message);
            }

        }




        public Int32 ConvertDaytoSeconds()
        {
            try
            {
                return IntervalDay * 24 * 60 * 60;
            }
            catch (Exception ex)
            {
                logger.Error("ConvertDaytoSeconds Function");
                logger.Error(ex.Message);
                throw;
            }
        }

        public Int32 ConvertDaytoSeconds(int intervalDay)
        {
            try
            {
                return intervalDay * 24 * 60 * 60;
            }
            catch (Exception ex)
            {
                logger.Error("ConvertDaytoSeconds Function");
                logger.Error(ex.Message);
                throw;
            }
        }

        public Int32 ConvertHourtoSeconds()
        {
            try
            {
                return IntervalHour * 60 * 60;
            }
            catch (Exception ex)
            {
                logger.Error("ConvertHourtoSeconds Function");
                logger.Error(ex.Message);
                throw;
            }
        }

        public Int32 ConvertHourtoSeconds(int intervalHour)
        {
            try
            {
                return intervalHour * 60 * 60;
            }
            catch (Exception ex)
            {
                logger.Error("ConvertHourtoSeconds Function");
                logger.Error(ex.Message);
                throw;
            }
        }

        public Int32 ConvertMinutestoSeconds()
        {
            try
            {
                return IntervalMinute * 60;
            }
            catch (Exception ex)
            {
                logger.Error("ConvertMinutestoSeconds Function");
                logger.Error(ex.Message);
                throw;
            }
        }

        public Int32 ConvertMinutestoSeconds(int intervalMinutes)
        {
            try
            {
                return intervalMinutes * 60;
            }
            catch (Exception ex)
            {
                logger.Error("ConvertMinutestoSeconds Function");
                logger.Error(ex.Message);
                throw;
            }
        }

        public Int32 GetIntervalinSeconds()
        {
            try
            {
                return ConvertDaytoSeconds() + ConvertHourtoSeconds() + ConvertMinutestoSeconds() + IntervalSecond;
            }
            catch (Exception ex)
            {
                logger.Error("GetInterval Function");
                logger.Error(ex.Message);
                throw;
            }
        }

        public Int32 GetIntervalinSeconds(QuartzParam quartzParam)
        {
            try
            {
                return ConvertDaytoSeconds(quartzParam.IntervalDay) + ConvertHourtoSeconds(quartzParam.IntervalHour) +
                       ConvertMinutestoSeconds(quartzParam.IntervalMinute) + quartzParam.IntervalSecond;
            }
            catch (Exception ex)
            {
                logger.Error("GetInterval Function");
                logger.Error(ex.Message);
                throw;
            }
        }

        //public bool GetIsDaily()
        //{
        //    try
        //    {
        //        return bool.Parse(ConfigurationManager.AppSettings["isDaily"]);
        //    }
        //    catch (Exception ex)
        //    {
        //        logger.Error("GetIsDaily function");
        //        logger.Error(ex.Message);
        //        logger.Error("Error getting isDaily Boolean, returning false as default");
        //        return false;
        //    }
        //}

        //public bool GetIsMonthly()
        //{
        //    try
        //    {
        //        return bool.Parse(ConfigurationManager.AppSettings["isMonthly"]);
        //    }
        //    catch (Exception ex)
        //    {
        //        logger.Error("GetIsMonthly function");
        //        logger.Error(ex.Message);
        //        logger.Error("Error getting GetIsMonthly Boolean, returning false as default");
        //        return false;
        //    }
        //}

        //public bool GetIsWeekly()
        //{
        //    try
        //    {
        //        return bool.Parse(ConfigurationManager.AppSettings["isWeekly"]);
        //    }
        //    catch (Exception ex)
        //    {
        //        logger.Error("GetIsWeekly function");
        //        logger.Error(ex.Message);
        //        logger.Error("Error getting GetIsWeekly Boolean, returning false as default");
        //        return false;
        //    }
        //}

        public DayOfWeek GetDayofWeek()
        {
            try
            {
                return (DayOfWeek)Enum.Parse(typeof(DayOfWeek), ConfigurationManager.AppSettings["DayOfWeek"], true);
            }
            catch (Exception ex)
            {
                logger.Error("GetDayofWeek function");
                logger.Error(ex.Message);
                logger.Error("Error getting GetDayofWeek Boolean, returning monday as default");
                return DayOfWeek.Monday;
            }
        }

        public DayOfWeek GetDayofWeek(String dayofWeek)
        {
            try
            {
                return (DayOfWeek)Enum.Parse(typeof(DayOfWeek), dayofWeek, true);
            }
            catch (Exception ex)
            {
                logger.Error("GetDayofWeek function");
                logger.Error(ex.Message);
                logger.Error("Error getting GetDayofWeek Boolean, returning monday as default");
                return DayOfWeek.Monday;
            }
        }

        //public int GetDayofMonth()
        //{
        //    try
        //    {
        //        return Int32.Parse(ConfigurationManager.AppSettings["DayOfMonth"]);
        //    }
        //    catch (Exception ex)
        //    {
        //        logger.Error("GetDayofMonth function");
        //        logger.Error(ex.Message);
        //        logger.Error("Error getting GetDayofMonth Boolean, returning 1 as default");
        //        return 1;
        //    }
        //}

        //public Int32 GetHour()
        //{
        //    try
        //    {
        //        return Int32.Parse(ConfigurationManager.AppSettings["StartHour"]);
        //    }
        //    catch (Exception ex)
        //    {
        //        logger.Error("GetHour Function");
        //        logger.Error(ex.Message);
        //        logger.Error("Error getting hour, returning 12 as default");
        //        return 12;
        //    }
        //}

        //public Int32 GetMinutes()
        //{
        //    try
        //    {
        //        return Int32.Parse(ConfigurationManager.AppSettings["StartMinutes"]);
        //    }
        //    catch (Exception ex)
        //    {
        //        logger.Error("GetHour Function");
        //        logger.Error(ex.Message);
        //        logger.Error("Error getting minutes, returning 0 as default");
        //        return 0;
        //    }
        //}

        public string ExecuteidleEmail()
        {
            String ErrorString;
            try
            {
                this.ConnectSQLServer();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = conSql;
                cmd.CommandText = "sp_Email";


                logger.Debug(cmd.CommandText);

                cmd.CommandType = CommandType.StoredProcedure;

                cmd.ExecuteNonQuery();

                this.CloseSqlServer();
                ErrorString = "Success";
                return ErrorString;
            }
            catch (Exception e)
            {
                logger.Error("ExecuteidleEmail Function");
                logger.Error(e.Message);
                this.CloseSqlServer();
                return null;
            }
        }


    }
}
