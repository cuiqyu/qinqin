package com.littlebayo.project.monitor.job.util;

import org.quartz.DisallowConcurrentExecution;
import org.quartz.JobExecutionContext;
import com.littlebayo.project.monitor.job.domain.Job;

/**
 * 定时任务处理（禁止并发执行）
 * 
 * @author littlebayo
 *
 */
@DisallowConcurrentExecution
public class QuartzDisallowConcurrentExecution extends AbstractQuartzJob
{
    @Override
    protected void doExecute(JobExecutionContext context, Job job) throws Exception
    {
        JobInvokeUtil.invokeMethod(job);
    }
}
