package com.littlebayo.project.qinq.pengda.domain;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;
import java.util.Map;

/**
 * 钉钉考勤报表月统计
 *
 * @author cuiqiongyu
 * @date 2020/10/27 10:33
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class DingdingMonthlyStatistics implements Serializable {

   private static final long serialVersionUID = -1635819533374757774L;

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
     * userId
     */
    private String userId;
    /**
     * 出勤天数
     */
    private String chuqinTianshu;
    /**
     * 休息天数
     */
    private String xiuxiTianshu;
    /**
     * 工作时长
     */
    private String gongzuoShichang;
    /**
     * 迟到次数
     */
    private String chidaoCishu;
    /**
     * 迟到时长
     */
    private String chidaoShichang;
    /**
     * 严重迟到次数
     */
    private String yanzhongChidaoCishu;
    /**
     * 严重迟到时长
     */
    private String yanzhongChidaoShichang;
    /**
     * 旷工迟到天数
     */
    private String kuanggongChidaoTianshu;
    /**
     * 早退次数
     */
    private String zaotuiCishu;
    /**
     * 早退时长
     */
    private String zaotuiShichang;
    /**
     * 上班缺卡次数
     */
    private String shangbanQuekaCishu;
    /**
     * 下班缺卡次数
     */
    private String xiabanQuekaCishu;
    /**
     * 旷工天数
     */
    private String kuangggongTianshu;
    /**
     * 外出（时）
     */
    private String waichu;
    /**
     * 丧假（天）
     */
    private String sangjia;
    /**
     * 事假（时）
     */
    private String shijia;
    /**
     * 病假（时）
     */
    private String bingjia;
    /**
     * 调休（时）
     */
    private String tiaoxiu;
    /**
     * 每日打卡总结
     */
    private Map<Integer, String> everydaySummary;

}
