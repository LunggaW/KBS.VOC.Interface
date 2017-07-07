using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using KBS.RANCH.VOCOLLECT.INTERFACE.MODEL;
using Quartz;
using Quartz.Impl;

namespace KBS.RANCH.VOC.INTERFACE.IDLEEMAIL
{
    public partial class ServiceIdleEmail : ServiceBase
    {
        private static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        private IdleEmail VocFunction = new IdleEmail();

        public ServiceIdleEmail()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            try
            {

                QuartzParam quartzParam = new QuartzParam();

                quartzParam = VocFunction.getQuartzParam();

                // construct a scheduler factory
                ISchedulerFactory schedFact = new StdSchedulerFactory();

                // get a scheduler
                IScheduler sched = schedFact.GetScheduler();
                sched.Start();

                // define the job and tie it to our HelloJob class
                IJobDetail job = JobBuilder.Create<HelloJob>()
                    .WithIdentity("myJob", "group1")
                    .Build();

                // Trigger the job to run now, and then every 40 seconds
                //ITrigger trigger = TriggerBuilder.Create()
                //  .WithIdentity("myTrigger", "group1")
                //  .StartNow()
                //  .WithSimpleSchedule(x => x
                //      .WithIntervalInSeconds(40)
                //      .RepeatForever())
                //  .Build();

                ITrigger trigger;


                //holiday calendar
                //HolidayCalendar cal = new HolidayCalendar();
                //cal.AddExcludedDate(DateTime.Now.AddDays(1));

                //sched.AddCalendar("myHolidays", cal, false,false);

                //if (VocFunction.GetIsDaily())
                //{
                //    logger.Debug("Start Daily");
                //    trigger = TriggerBuilder.Create()
                //        .WithIdentity("myTrigger")
                //        .WithSchedule(CronScheduleBuilder.DailyAtHourAndMinute(VocFunction.GetHour(),
                //            VocFunction.GetMinutes())) // execute job daily at
                //        //.ModifiedByCalendar("myHolidays") // but not on holidays
                //        .Build();
                //}
                //else if (VocFunction.GetIsWeekly())
                //{
                //    logger.Debug("Start Weekly");
                //    logger.Debug("Day of week is : " + VocFunction.GetDayofWeek());
                //    trigger = TriggerBuilder.Create()
                //        .WithIdentity("myTrigger")
                //        .WithSchedule(CronScheduleBuilder.WeeklyOnDayAndHourAndMinute(VocFunction.GetDayofWeek(),
                //            VocFunction.GetHour(),
                //            VocFunction.GetMinutes())) // execute job daily at
                //        //.ModifiedByCalendar("myHolidays") // but not on holidays
                //        .Build();
                //}
                //else if (VocFunction.GetIsMonthly())
                //{
                //    logger.Debug("Start Monthly");
                //    logger.Debug("Day of month is : " + VocFunction.GetDayofMonth());
                //    trigger = TriggerBuilder.Create()
                //        .WithIdentity("myTrigger")
                //        .WithSchedule(CronScheduleBuilder.MonthlyOnDayAndHourAndMinute(VocFunction.GetDayofMonth(),
                //            VocFunction.GetHour(),
                //            VocFunction.GetMinutes())) // execute job daily at
                //        //.ModifiedByCalendar("myHolidays") // but not on holidays
                //        .Build();
                //}
                //else
                //{
                //    trigger = TriggerBuilder.Create()
                //        .WithIdentity("myTrigger", "group1")
                //        .WithSimpleSchedule(x => x
                //            .WithIntervalInSeconds(VocFunction.GetIntervalinSeconds())
                //            .RepeatForever())
                //        .Build();
                //}


                if (quartzParam.IsDaily)
                {
                    logger.Debug("Start Daily");
                    trigger = TriggerBuilder.Create()
                        .WithIdentity("myTrigger")
                        .WithSchedule(CronScheduleBuilder.DailyAtHourAndMinute(quartzParam.StartHour,
                            quartzParam.StartMinute)) // execute job daily at
                        //.ModifiedByCalendar("myHolidays") // but not on holidays
                        .Build();
                }
                else if (quartzParam.IsWeekly)
                {
                    logger.Debug("Start Weekly");
                    logger.Debug("Day of week is : " + VocFunction.GetDayofWeek());
                    trigger = TriggerBuilder.Create()
                        .WithIdentity("myTrigger")
                        .WithSchedule(CronScheduleBuilder.WeeklyOnDayAndHourAndMinute(VocFunction.GetDayofWeek(quartzParam.DayOfWeek),
                            quartzParam.StartHour,
                            quartzParam.StartMinute)) // execute job daily at
                        //.ModifiedByCalendar("myHolidays") // but not on holidays
                        .Build();
                }
                else if (quartzParam.IsMonthly)
                {
                    logger.Debug("Start Monthly");
                    logger.Debug("Day of month is : " + quartzParam.DayOfMonth);
                    trigger = TriggerBuilder.Create()
                        .WithIdentity("myTrigger")
                        .WithSchedule(CronScheduleBuilder.MonthlyOnDayAndHourAndMinute(quartzParam.DayOfMonth,
                            quartzParam.StartHour,
                            quartzParam.StartMinute)) // execute job daily at
                        //.ModifiedByCalendar("myHolidays") // but not on holidays
                        .Build();
                }
                else
                {
                    trigger = TriggerBuilder.Create()
                        .WithIdentity("myTrigger", "group1")
                        .WithSimpleSchedule(x => x
                            .WithIntervalInSeconds(VocFunction.GetIntervalinSeconds(quartzParam))
                            .RepeatForever())
                        .Build();
                }




                sched.ScheduleJob(job, trigger);
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                logger.Error(ex.InnerException);
                throw;
            }
        }

        protected override void OnStop()
        {
        }

        public class HelloJob : IJob
        {
            private static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
            private IdleEmail VocFunction = new IdleEmail();


            public void Execute(IJobExecutionContext context)
            {
                logger.Debug("Start Job");
                StartJob();
            }

            public void StartJob()
            {
                // TODO: Insert monitoring activities here.
                //eventLog1.WriteEntry("Monitoring the System", EventLogEntryType.Information, eventId++);StringBuilder sb = new StringBuilder(); 
                //sales_int();
                //CSMV3Function.SelectSalesSql();
                //CSMV3Function.ExecExportDeliveryNote();
                //CSMV3Function.ExecExportBarcodeMaster();
                //CSMV3Function.ExecExportSalesPriceMaster();
                //CSMV3Function.ExecExportItemMaster();
                //CSMV3Function.ExecExportStoreMaster();
                ExecuteInterface();
            }

            public void ExecuteInterface()
            {
                try
                {
                    VocFunction.ExecuteidleEmail();
                }
                catch (Exception ex)
                {
                    logger.Error("Messsage : " + ex.Message);
                    logger.Error("Inner Exception : " + ex.InnerException);
                    throw;
                }
            }
        }
    }
}
