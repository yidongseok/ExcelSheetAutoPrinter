using log4net;
using Quartz;
using System;
using System.Reflection;
using System.Threading.Tasks;

namespace ExcelSheetAutoPrinter
{
	internal class TestJob : IJob
	{
		private static readonly ILog logger = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

		public async Task Execute(IJobExecutionContext context)
		{
			try
			{
				/*// Log the job execution
				logger.Info("TestJob is executing at: " + DateTime.Now);
				// Simulate some work
				await Task.Delay(1000); // Simulate a delay of 1 second
				// Log completion
				logger.Info("TestJob completed at: " + DateTime.Now);*/

				logger.Warn("수행할 작업 추가");
				await Task.Delay(1000); // Simulate a delay of 1 second
			}
			catch (Exception ex)
			{
				logger.Error(ex);
			}
		}
	}
}
