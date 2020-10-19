package com.littlebayo.project.qinq.pengda.domain;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;

/**
 * 钉钉打开记录
 *
 * @author cuiqiongyu
 * @date 2020/10/19 15:39
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class DingdingPunchInRecord implements Serializable {

    private static final long serialVersionUID = -2405560318146708564L;

    /**
     * userId
     */
    private String userId;
    /**
     * 日期
     */
    private String dateTime;
    /**
     * 班次
     */
    private String frequency;

    /**
     * 上班一打卡时间
     */
    private String goToWorkClockInTime1;
    /**
     * 上班一打卡结果
     */
    private String goToWorkClockoutResults1;
    /**
     * 下班一打卡时间
     */
    private String goOffWorkClockInTime1;
    /**
     * 下班一打卡结果
     */
    private String goOffWorkClockoutResults1;
    /**
     * 上班二打卡时间
     */
    private String goToWorkClockInTime2;
    /**
     * 上班二打卡结果
     */
    private String goToWorkClockoutResults2;
    /**
     * 下班二打卡时间
     */
    private String goOffWorkClockInTime2;
    /**
     * 下班二打卡结果
     */
    private String goOffWorkClockoutResults2;
    /**
     * 关联的审批单
     */
    private String approvalForm;

}
