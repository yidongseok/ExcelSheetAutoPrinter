using System;

namespace ExcelSheetAutoPrinter
{
	internal class JobInfo
	{
		/// <summary>
        /// Job Key
        /// </summary>
        public string Key
        {
            get;
            set;
        }

        /// <summary>
        /// Corn 표현식
        /// </summary>
        public string CronExpression
        {
            get;
            set;
        }

        /// <summary>
        /// Job 시작 시간
        /// </summary>
        public DateTime StartTime
        {
            get;
            set;
        }

        /// <summary>
        /// Job 종료 시간
        /// </summary>
        public DateTime EndTime
        {
            get;
            set;
        }
	}
}
