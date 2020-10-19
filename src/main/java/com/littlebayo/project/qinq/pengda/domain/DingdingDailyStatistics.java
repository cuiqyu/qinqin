package com.littlebayo.project.qinq.pengda.domain;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;
import java.util.List;

/**
 * 钉钉考勤报表每日统计格式
 *
 * @authocuiqiongyu
 * @date 2020/10/19 15:12
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class DingdingDailyStatistics implements Serializable{

    private static final long serialVersionUID = 8905036049308382862L;

    /**
     * 姓名
     */
    private String name;
    /**
     * 考勤组
     */
    private String attendanceContent;
    /**
     * 部门名称
     */
    private String deptName;
    /**
     * 工号
     */
    private String jobNumber;
    /**
     * 职位
     */
    private String jobName;
    /**
     * userId
     */
    private String userId;
    /**
     * 打卡时间记录
     */
    private List<DingdingPunchInRecord> punchInRecords;

}
