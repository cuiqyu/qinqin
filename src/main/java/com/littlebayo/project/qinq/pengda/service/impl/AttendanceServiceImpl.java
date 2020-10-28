package com.littlebayo.project.qinq.pengda.service.impl;

import com.google.common.collect.Lists;
import com.littlebayo.common.exception.BusinessException;
import com.littlebayo.common.utils.StringUtils;
import com.littlebayo.framework.config.RuoYiConfig;
import com.littlebayo.framework.web.domain.AjaxResult;
import com.littlebayo.project.qinq.pengda.domain.DingdingDailyStatistics;
import com.littlebayo.project.qinq.pengda.domain.DingdingMonthlyStatistics;
import com.littlebayo.project.qinq.pengda.domain.DingdingPunchInRecord;
import com.littlebayo.project.qinq.pengda.service.AttendanceService;
import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.*;
import java.util.stream.Collectors;

/**
 * 考勤服务
 *
 * @author cuiqiongyu
 * @date 2020/10/19 14:53
 */
@Service
public class AttendanceServiceImpl implements AttendanceService {

    private static final Logger logger = LoggerFactory.getLogger(AttendanceServiceImpl.class);

    /**
     * 上传报表所支持的格式
     */
    private static final List<String> ATTENDANCE_FILE_TYPE_SUFFIX = Lists.newArrayList(".xls", ".xlsx");
    /**
     * 需要解析的钉钉报表的表名
     */
    private static final String ATTENDANCE_DAILYSTATISTICSDATA_NAME = "每日统计";
    private static final String ATTENDANCE_MONTHLYSTATISTICSDATA_NAME = "月度汇总";
    /**
     * 生成的考勤报表添加的文件名后缀
     */
    private static final String ATTENDANCE_REPORT_SUFFIX = "_qinq";
    /**
     * 每日统计表解析的标题
     */
    private static String clockInTimeAndDate = "每日统计表 统计日期：2020-09-01 至 2020-09-30";
    /**
     * 员工大夜班补贴跟加班补贴次数
     */
    private Map<String, Integer> workOvertimeMap = new HashMap<>();
    private Map<String, Integer> nightShiftMap = new HashMap<>();

    /**
     * 导入钉钉文件
     *
     * @param file 上传的文件
     * @return 返回结果
     */
    @Override
    public AjaxResult importDingdingFile(MultipartFile file) throws Exception {
        if (null == file || file.isEmpty()) {
            logger.error("钉钉考勤文件上传失败，文件内容不能为空！");
            return AjaxResult.error("文件内容不能为空!");
        }

        // 获取文件上传名称
        String filename = file.getOriginalFilename();
        // 获取文件后缀
        String suffix = filename.substring(filename.lastIndexOf("."));
        // 文件类型检查
        if (!ATTENDANCE_FILE_TYPE_SUFFIX.contains(suffix)) {
            logger.error("钉钉考勤文件上传失败，文件类型不正确，只支持后缀为" + ATTENDANCE_FILE_TYPE_SUFFIX + "的文件");
            return AjaxResult.error("文件类型不正确，只支持后缀为" + ATTENDANCE_FILE_TYPE_SUFFIX + "的文件");
        }

        // 获取文件的真实名称，去除文件后缀的
        String fileRealName = filename.substring(0, filename.lastIndexOf("."));
        // 开始解析第一个表格文件
        List<DingdingMonthlyStatistics> dingdingMonthlyStatistics = parseDingdingExcelMonthlyStatisticsData(file.getInputStream());
        // 开始解析第二个表格文件
        List<DingdingDailyStatistics> dingdingDailyStatistics = parseDingdingExcelDailyStatisticsData(file.getInputStream());

        // -------开始分析并导出-------
        // 1. 创建工作簿对象
        Workbook wb = new SXSSFWorkbook(500);
        wb.createSheet(getMonthStr() + "月考勤表统计");
        wb.createSheet(getMonthStr() + "月打卡时间");

        // 2. 生成x月打卡时间【表二】【因为生成表一的时候需要用到一些表二的数据，所以先生成表二】
        wb = generatorXmonthClockInTime(dingdingDailyStatistics, wb);

        // 3. 生成x月考勤表统计【表一】
        wb = generatorXmonthCheckWorkAttendance(dingdingDailyStatistics, dingdingMonthlyStatistics, wb);

        // 4. 输出工作表
        String fileName = fileRealName + ATTENDANCE_REPORT_SUFFIX + suffix;
        OutputStream out = new FileOutputStream(getAbsoluteFile(fileName));
        wb.write(out);
        return AjaxResult.success(fileName);
    }

    /**
     * 解析钉钉考勤报表每月统计（第一张表格）
     *
     * @param inputStream 文件输入流
     * @return
     */
    private List<DingdingMonthlyStatistics> parseDingdingExcelMonthlyStatisticsData(InputStream inputStream) throws IOException, InvalidFormatException {
        // 工作簿对象
        Workbook wb = WorkbookFactory.create(inputStream);
        // 获取【月度汇总】表格
        Sheet sheet = wb.getSheet(ATTENDANCE_MONTHLYSTATISTICSDATA_NAME);
        if (null == sheet) {
            logger.error("解析钉钉考勤报表失败，未找到报表文件中的[月度汇总]表格");
            throw new BusinessException("未找到报表文件中的[月度汇总]表格");
        }

        // 月度汇总数据内容
        List<DingdingMonthlyStatistics> dataList = new ArrayList<>();
        // 获取总行数
        int rows = sheet.getPhysicalNumberOfRows();
        if (rows <= 4) {
            return dataList;
        }

        // 开始统计的行数 从第五行开始统计
        for (int i = 4; i < rows; i++) {
            Row row = sheet.getRow(i);
            int j = 0;
            DingdingMonthlyStatistics monthlyStatistics = new DingdingMonthlyStatistics(
                    getCellStringValue(row.getCell(j++)), getCellStringValue(row.getCell(j++)), getCellStringValue(row.getCell(j++)),
                    getCellStringValue(row.getCell(j++)), getCellStringValue(row.getCell(j++)), getCellStringValue(row.getCell(j++)),
                    getCellStringValue(row.getCell(j++)), getCellStringValue(row.getCell(j++)), getCellStringValue(row.getCell(j++)),
                    getCellStringValue(row.getCell(j++)), getCellStringValue(row.getCell(j++)), getCellStringValue(row.getCell(j++)),
                    getCellStringValue(row.getCell(j++)), getCellStringValue(row.getCell(j++)), getCellStringValue(row.getCell(j++)),
                    getCellStringValue(row.getCell(j++)), getCellStringValue(row.getCell(j++)), getCellStringValue(row.getCell(j++)),
                    getCellStringValue(row.getCell(j++)), getCellStringValue(row.getCell(j++)), getCellStringValue(row.getCell(j++)),
                    getCellStringValue(row.getCell(j++)), getCellStringValue(row.getCell(j++)), null
            );

            HashMap<Integer, String> hashMap = new HashMap<>();
            // 假设每个月按照最大的天数 31天来算
            for (int m = 1; m <= 31; m++) {
                hashMap.put(m, getCellStringValue(row.getCell(j++)));
            }
            monthlyStatistics.setEverydaySummary(hashMap);
            dataList.add(monthlyStatistics);
        }

        return dataList;
    }

    /**
     * 解析钉钉考勤报表每日统计（第二张表格）
     *
     * @param inputStream 文件输入流
     */
    private List<DingdingDailyStatistics> parseDingdingExcelDailyStatisticsData(InputStream inputStream) throws IOException, InvalidFormatException {
        // 工作簿对象
        Workbook wb = WorkbookFactory.create(inputStream);
        // 获取【每日统计】表格
        Sheet sheet = wb.getSheet(ATTENDANCE_DAILYSTATISTICSDATA_NAME);
        if (null == sheet) {
            logger.error("解析钉钉考勤报表失败，未找到报表文件中的[每日统计]表格");
            throw new BusinessException("未找到报表文件中的[每日统计]表格");
        }

        // 每日统计数据内容
        List<DingdingDailyStatistics> dataList = new ArrayList<>();
        // 获取总行数
        int rows = sheet.getPhysicalNumberOfRows();
        if (rows <= 4) {
            return dataList;
        }
        // 读取每日统计的打卡时间日期
        clockInTimeAndDate = getCellStringValue(sheet.getRow(0).getCell(0));

        // 打卡时间记录
        List<DingdingPunchInRecord> punchInRecords = new ArrayList<>(2000);
        // 已经记录的人员
        Set<String> userIdSet = new HashSet<>(70);

        // 开始统计的行数 从第五行开始统计
        for (int i = 4; i < rows; i++) {
            Row row = sheet.getRow(i);
            String userId = getCellStringValue(row.getCell(5));
            if (!userIdSet.contains(userId)) {
                DingdingDailyStatistics dd = new DingdingDailyStatistics();
                // 姓名
                dd.setName(getCellStringValue(row.getCell(0)));
                // 考勤组
                dd.setAttendanceContent(getCellStringValue(row.getCell(1)));
                // 部门名称
                dd.setDeptName(getCellStringValue(row.getCell(2)));
                // 工号
                dd.setJobNumber(getCellStringValue(row.getCell(3)));
                // 职位
                dd.setJobName(getCellStringValue(row.getCell(4)));
                // userId
                dd.setUserId(userId);
                // 添加到解析结果中
                dataList.add(dd);
                userIdSet.add(userId);
            }

            DingdingPunchInRecord record = new DingdingPunchInRecord(
                    userId,
                    getCellStringValue(row.getCell(6)),
                    getCellStringValue(row.getCell(7)),
                    getCellStringValue(row.getCell(8)),
                    getCellStringValue(row.getCell(9)),
                    getCellStringValue(row.getCell(10)),
                    getCellStringValue(row.getCell(11)),
                    getCellStringValue(row.getCell(12)),
                    getCellStringValue(row.getCell(13)),
                    getCellStringValue(row.getCell(14)),
                    getCellStringValue(row.getCell(15)),
                    getCellStringValue(row.getCell(16))
            );
            punchInRecords.add(record);
        }

        // 将打卡记录根据userId转成HashMap结构
        Map<String, List<DingdingPunchInRecord>> punchInRecordMap = punchInRecords.stream()
                .collect(Collectors.toMap(d -> d.getUserId(), d -> Lists.newArrayList(d), (l1, l2) -> {
                    l1.addAll(l2);
                    return l1;
                }));

        // 将打卡记录塞入每日统计结果中
        if (CollectionUtils.isNotEmpty(dataList)) {
            for (DingdingDailyStatistics statistic : dataList) {
                if (punchInRecordMap.containsKey(statistic.getUserId())) {
                    statistic.setPunchInRecords(punchInRecordMap.get(statistic.getUserId()));
                }
            }
        }

        return dataList;
    }

    /**
     * TODO 生成x月考勤表统计【表一】
     *
     * @param dingdingDailyStatistics   用户的打卡记录
     * @param dingdingMonthlyStatistics 用户的打卡月统计
     * @param wb                        输出的excel
     * @return 输出excel
     */
    private Workbook generatorXmonthCheckWorkAttendance(List<DingdingDailyStatistics> dingdingDailyStatistics,
                                                        List<DingdingMonthlyStatistics> dingdingMonthlyStatistics, Workbook wb) {
        if (null == wb || CollectionUtils.isEmpty(dingdingDailyStatistics)) {
            return wb;
        }
        Sheet sheet = wb.getSheet(getMonthStr() + "月考勤表统计");
        Drawing draw = sheet.createDrawingPatriarch();
        // 设置列宽
        List<Integer> columnWidthList = Lists.newArrayList(5, 14, 5, 8, 6, 6, 6, 3, 3, 3, 4, 3, 3, 3, 3, 5, 3, 3, 3, 3, 3, 3, 3,
                3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3);
        for (int i = 0; i < columnWidthList.size(); i++) {
            sheet.setColumnWidth(i, columnWidthList.get(i) * 256);
        }

        // 获取日期列表
        List<DingdingPunchInRecord> punchInRecords = dingdingDailyStatistics.get(0).getPunchInRecords();
        // 行标
        int rowIndex = 0;
        /**
         * 开始填充数据
         */

        /**
         * 1.设置第一行的标题
         */
        Row row0 = sheet.createRow(rowIndex++);
        row0.setHeightInPoints(32.25f);
        Cell row0cell0 = row0.createCell(0, CellType.STRING);
        row0cell0.setCellValue("杭州芃达网络科技有限公司考勤表");
        // 设置第一行标题的单元格格式
        CellStyle titleCellStyle = wb.createCellStyle();
        titleCellStyle.setAlignment(HorizontalAlignment.CENTER);
        titleCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCellStyle.setBorderBottom(BorderStyle.THIN);
        titleCellStyle.setBorderTop(BorderStyle.THIN);
        titleCellStyle.setBorderLeft(BorderStyle.THIN);
        titleCellStyle.setBorderRight(BorderStyle.THIN);
        Font font = wb.createFont();
        font.setFontName("微软雅黑");
        font.setFontHeightInPoints((short) 24);
        titleCellStyle.setFont(font);
        row0cell0.setCellStyle(titleCellStyle);
        CellRangeAddress region = new CellRangeAddress(0, 0, 0, punchInRecords.size() + 15);
        sheet.addMergedRegion(region);

        /**
         * 2. 设置第二行的标题
         */
        Row row1 = sheet.createRow(rowIndex++);
        row1.setHeightInPoints(16.5f);
        Cell row1cell0 = row1.createCell(0, CellType.STRING);
        row1cell0.setCellValue(getMonthStr() + "月份考勤明细总汇表");
        // 设置第一行标题的单元格格式
        CellStyle titleCellStyle2 = wb.createCellStyle();
        titleCellStyle2.setAlignment(HorizontalAlignment.CENTER);
        titleCellStyle2.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCellStyle2.setWrapText(true);
        titleCellStyle2.setBorderBottom(BorderStyle.THIN);
        titleCellStyle2.setBorderTop(BorderStyle.THIN);
        titleCellStyle2.setBorderLeft(BorderStyle.THIN);
        titleCellStyle2.setBorderRight(BorderStyle.THIN);
        Font font1 = wb.createFont();
        font1.setFontName("微软雅黑");
        font1.setFontHeightInPoints((short) 11);
        titleCellStyle2.setFont(font1);
        row1cell0.setCellStyle(titleCellStyle2);
        CellRangeAddress region1 = new CellRangeAddress(1, 1, 0, punchInRecords.size() + 15);
        sheet.addMergedRegion(region1);

        /**
         * 3. 设置第三行的表头
         */
        Row row2 = sheet.createRow(rowIndex++);
        row2.setHeightInPoints(16.5f);
        // 获取日期列表
        Map<String, String> datatimeStrMap = punchInRecords.stream().map(p -> p.getDateTime()).collect(Collectors.toMap(
                p -> {
                    p = p.substring(p.lastIndexOf("-") + 1, p.indexOf("星期"));
                    return p.charAt(0) == '0' ? p.substring(1).trim() : p.trim();
                },
                p -> p.substring(p.indexOf("星期")),
                (l1, l2) -> l2));
        int row2cellIndex = 0;
        List<String> titleNameList = Lists.newArrayList("序号", "姓名", "应出勤天数", "实际出勤天数", "缺勤", "平时加班天数",
                "事假天数", "产假天数", "丧假天数", "婚假天数", "病假天数", "迟到早退次数", "漏打卡次数", "大夜班补贴", "加班补贴", "日期时间");
        for (int i = 0; i < titleNameList.size(); i++) {
            Cell row2Celln = row2.createCell(row2cellIndex++, CellType.STRING);
            row2Celln.setCellValue(titleNameList.get(i));
            row2Celln.setCellStyle(titleCellStyle2);
        }
        List<String> list = datatimeStrMap.keySet().stream().sorted(Comparator.comparing(Integer::valueOf)).collect(Collectors.toList());
        for (String key : list) {
            Cell cell = row2.createCell(row2cellIndex++, CellType.STRING);
            cell.setCellValue(key);
            cell.setCellStyle(titleCellStyle2);
        }
        // 合并单元格
        for (int i = 0; i < 16; i++) {
            CellRangeAddress regionN = new CellRangeAddress(2, 3, i, i);
            sheet.addMergedRegion(regionN);
        }
        /**
         * 设置第四行内容
         */
        Row row3 = sheet.createRow(rowIndex++);
        row3.setHeightInPoints(49.5f);
        int row3cellIndex = 16;
        for (String key : list) {
            Cell row3Celln = row3.createCell(row3cellIndex++, CellType.STRING);
            row3Celln.setCellValue(datatimeStrMap.get(key));
            row3Celln.setCellStyle(titleCellStyle2);
        }
        // 将每月统计数据按照userId来分组
        Map<String, DingdingMonthlyStatistics> monthlyStatisticsMap = dingdingMonthlyStatistics.stream()
                .collect(Collectors.toMap(m -> m.getUserId(), m -> m, (m1, m2) -> m2));

        // 黄色背景时间样式
        CellStyle yellowCellStyle = wb.createCellStyle();
        yellowCellStyle.cloneStyleFrom(titleCellStyle2);
        yellowCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        yellowCellStyle.setFillForegroundColor(HSSFColor.YELLOW.index);
        // 未正确实现的单元格，需要手工处理
        CellStyle todoCellStyle = wb.createCellStyle();
        todoCellStyle.cloneStyleFrom(titleCellStyle2);
        todoCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        todoCellStyle.setFillForegroundColor(HSSFColor.ORANGE.index);

        // 开始填充数据
        for (int i = 0; i < dingdingDailyStatistics.size(); i++) {
            DingdingDailyStatistics statistics = dingdingDailyStatistics.get(i);
            String userId = statistics.getUserId();
            DingdingMonthlyStatistics monthlyStatistics = monthlyStatisticsMap.get(userId);
            if (null == monthlyStatistics) {
                monthlyStatistics = new DingdingMonthlyStatistics();
            }
            // 去掉王金玉和王和玉
            if (userId.equals("066937470829612195") || userId.equals("manager3394")) {
                continue;
            }

            Row rown1 = sheet.createRow(rowIndex++);
            int rown1cellIndex = 0;
            Row rown2 = sheet.createRow(rowIndex++);
            int rown2cellIndex = 0;

            /**
             * 设置合并两行的，前0-15列的信息
             */
            Integer workOvertimeCount = workOvertimeMap.get(userId);
            Integer nightShiftCount = nightShiftMap.get(userId);
            List<String> dayContentList = Lists.newArrayList(String.valueOf(i + 1), statistics.getName(), "",
                    monthlyStatistics.getChuqinTianshu(), "", "", "", "", "", "", "", "", "",
                    String.valueOf(null == nightShiftCount ? 0 : nightShiftCount), String.valueOf(null == workOvertimeCount ? 0 : workOvertimeCount), "");
            for (String day : dayContentList) {
                Cell rown1Cell = rown1.createCell(rown1cellIndex++, CellType.STRING);
                rown1Cell.setCellValue(day);
                rown1Cell.setCellStyle(titleCellStyle2);
            }

            // 合并某些单元格的第一行第二行
            for (int j = 0; j < 16; j++) {
                rown2.createCell(rown2cellIndex++, CellType.STRING).setCellValue("");
                CellRangeAddress regionN12 = new CellRangeAddress(rowIndex - 2, rowIndex - 1, j, j);
                sheet.addMergedRegion(regionN12);
            }

            // TODO未实现的单元格
            for (Integer todo : Lists.newArrayList(2, 4, 5, 6, 7, 8, 9, 10, 11, 12)) {
                rown1.getCell(todo).setCellStyle(todoCellStyle);
                rown2.getCell(todo).setCellStyle(todoCellStyle);
            }

            /**
             * 开始设置第一行跟第二行的打卡时间
             * 上班"√"
             * 大夜班"√"
             * 加班"√"
             * 休息"●" 【黄色背景标志】
             * 事假"×"
             * 病假"△"
             * 旷工"○"
             * 迟到"★"【加批注】
             * 早退"▲"
             * 漏打卡"⊙"
             * 婚嫁"+"丧假"±"离职"＃"
             * 工伤生育假"※"
             */
            List<DingdingPunchInRecord> punchInRecords1 = statistics.getPunchInRecords();
            if (CollectionUtils.isNotEmpty(punchInRecords1)) {
                // 开始遍历打卡记录
                for (int j = 1; j <= punchInRecords1.size(); j++) {
                    Map<Integer, String> everydaySummary = monthlyStatistics.getEverydaySummary();
                    everydaySummary = (null != everydaySummary && !everydaySummary.isEmpty()) ? everydaySummary : new HashMap<>();
                    String jSummary = everydaySummary.get(j);
                    Cell rown1Cell = rown1.createCell(rown1cellIndex++, CellType.STRING);
                    Cell rown2Cell = rown2.createCell(rown2cellIndex++, CellType.STRING);
                    // 没有打卡记录说明【未入职、离职】
                    if (StringUtils.isEmpty(jSummary) || jSummary.equals("不在考勤组并打卡")) {
                        // 判断前一列，如果不为null，则该单元格置为#号 或者当前格是第16列，则也置为#号
                        if (rown1cellIndex == 17) {
                            rown1Cell.setCellValue("#");
                            rown1Cell.setCellStyle(titleCellStyle2);
                            rown2Cell.setCellValue("#");
                            rown2Cell.setCellStyle(titleCellStyle2);
                        } else {
                            Cell cell1n = rown1.getCell(rown1cellIndex - 1);
                            if (StringUtils.isEmpty(cell1n.getStringCellValue())) {
                                rown1Cell.setCellValue("#");
                                rown1Cell.setCellStyle(titleCellStyle2);
                            }
                            Cell cell2n = rown2.getCell(rown1cellIndex - 1);
                            if (StringUtils.isEmpty(cell2n.getStringCellValue())) {
                                rown2Cell.setCellValue("#");
                                rown2Cell.setCellStyle(titleCellStyle2);
                            }
                        }
                    }
                    // 表示有打卡记录内容的
                    else {
                        // 月打卡统计提示该天正常
                        if (jSummary.equals("正常")) {
                            // 则上午下午打卡都正常
                            rown1Cell.setCellValue("√");
                            rown1Cell.setCellStyle(titleCellStyle2);
                            rown2Cell.setCellValue("√");
                            rown2Cell.setCellStyle(titleCellStyle2);
                        }
                        // 月打卡改天休息或者旷工
                        else if (jSummary.equals("休息") || jSummary.equals("旷工")) {
                            // 则上午下午都休息
                            rown1Cell.setCellValue("●");
                            rown1Cell.setCellStyle(yellowCellStyle);
                            rown2Cell.setCellValue("●");
                            rown2Cell.setCellStyle(yellowCellStyle);
                        }
                        // 月打卡该天上班迟到几分钟
                        else if (jSummary.indexOf("上班迟到") > -1) {
                            // 则上午迟到几分钟，下午正常
                            rown1Cell.setCellValue("★");
                            rown1Cell.setCellStyle(titleCellStyle2);
                            // 批注
                            Comment comment = draw.createCellComment(new XSSFClientAnchor(0, 0, 0, 0,
                                    rown1cellIndex - 1, rowIndex - 2, rown1cellIndex - 1, rowIndex - 2));
                            XSSFRichTextString rtf = new XSSFRichTextString(jSummary.replace("上班", ""));
                            comment.setString(rtf);
                            rown1Cell.setCellComment(comment);
                            rown2Cell.setCellValue("√");
                            rown2Cell.setCellStyle(titleCellStyle2);
                        }
                        // 月打卡-上班缺卡
                        else if (jSummary.equals("上班缺卡")) {
                            // 则上午补卡下午打卡正常
                            rown1Cell.setCellValue("⊙");
                            rown1Cell.setCellStyle(titleCellStyle2);
                            rown2Cell.setCellValue("√");
                            rown2Cell.setCellStyle(titleCellStyle2);
                        }
                        // 月打卡-下班缺卡
                        else if (jSummary.equals("下班缺卡")) {
                            // 则下午补卡上午打卡正常
                            rown1Cell.setCellValue("√");
                            rown1Cell.setCellStyle(titleCellStyle2);
                            rown2Cell.setCellValue("⊙");
                            rown2Cell.setCellStyle(titleCellStyle2);
                        }
                        // 调休|事假|丧假
                        else if (jSummary.startsWith("调休") || jSummary.startsWith("事假") || jSummary.startsWith("丧假")) {
                            try {
                                // 获取最后的调休时长
                                String substring = jSummary.substring(jSummary.lastIndexOf(" ") + 1);
                                if (substring.endsWith("小时")) {
                                    String su = substring.substring(0, substring.length() - 2);
                                    Integer suInteger = Integer.valueOf(su);
                                    if (suInteger >= 8) {
                                        // 则上午下午都休息
                                        rown1Cell.setCellValue("●");
                                        rown1Cell.setCellStyle(yellowCellStyle);
                                        rown2Cell.setCellValue("●");
                                        rown2Cell.setCellStyle(yellowCellStyle);
                                        if (jSummary.startsWith("丧假")) {
                                            // 批注
                                            Comment comment = draw.createCellComment(new XSSFClientAnchor(0, 0, 0, 0,
                                                    rown1cellIndex - 1, rowIndex - 2, rown1cellIndex - 1, rowIndex - 2));
                                            XSSFRichTextString rtf = new XSSFRichTextString(jSummary);
                                            comment.setString(rtf);
                                            rown1Cell.setCellComment(comment);
                                            rown2Cell.setCellComment(comment);
                                        }
                                    }
                                } else if (substring.endsWith("天")) {
                                    String su = substring.substring(0, substring.length() - 1);
                                    Integer suInteger = Integer.valueOf(su);
                                    if (suInteger >= 1) {
                                        // 则上午下午都休息
                                        rown1Cell.setCellValue("●");
                                        rown1Cell.setCellStyle(yellowCellStyle);
                                        rown2Cell.setCellValue("●");
                                        rown2Cell.setCellStyle(yellowCellStyle);
                                        if (jSummary.startsWith("丧假")) {
                                            // 批注
                                            Comment comment = draw.createCellComment(new XSSFClientAnchor(0, 0, 0, 0,
                                                    rown1cellIndex - 1, rowIndex - 2, rown1cellIndex - 1, rowIndex - 2));
                                            XSSFRichTextString rtf = new XSSFRichTextString(jSummary);
                                            comment.setString(rtf);
                                            rown1Cell.setCellComment(comment);
                                            rown2Cell.setCellComment(comment);
                                        }
                                    }
                                }
                            } catch (Exception e) {
                                rown1Cell.setCellValue("");
                                rown1Cell.setCellStyle(todoCellStyle);
                                rown2Cell.setCellValue("");
                                rown2Cell.setCellStyle(todoCellStyle);
                                // 批注
                                Comment comment = draw.createCellComment(new XSSFClientAnchor(0, 0, 0, 0,
                                        rown1cellIndex - 1, rowIndex - 2, rown1cellIndex - 1, rowIndex - 2));
                                XSSFRichTextString rtf = new XSSFRichTextString(jSummary);
                                comment.setString(rtf);
                                rown1Cell.setCellComment(comment);
                                // 批注
                                Comment comment1 = draw.createCellComment(new XSSFClientAnchor(0, 0, 0, 0,
                                        rown1cellIndex - 1, rowIndex - 1, rown1cellIndex - 1, rowIndex - 1));
                                XSSFRichTextString rtf1 = new XSSFRichTextString(jSummary);
                                comment1.setString(rtf1);
                                rown2Cell.setCellComment(comment1);
                            }
                        } else {
                            rown1Cell.setCellValue("");
                            rown1Cell.setCellStyle(todoCellStyle);
                            rown2Cell.setCellValue("");
                            rown2Cell.setCellStyle(todoCellStyle);
                            // 批注
                            Comment comment = draw.createCellComment(new XSSFClientAnchor(0, 0, 0, 0,
                                    rown1cellIndex - 1, rowIndex - 2, rown1cellIndex - 1, rowIndex - 2));
                            XSSFRichTextString rtf = new XSSFRichTextString(jSummary);
                            comment.setString(rtf);
                            rown1Cell.setCellComment(comment);
                            // 批注
                            Comment comment1 = draw.createCellComment(new XSSFClientAnchor(0, 0, 0, 0,
                                    rown1cellIndex - 1, rowIndex - 1, rown1cellIndex - 1, rowIndex - 1));
                            XSSFRichTextString rtf1 = new XSSFRichTextString(jSummary);
                            comment1.setString(rtf1);
                            rown2Cell.setCellComment(comment1);
                        }
                    }
                }
            }
        }

        /**
         * 设置最后一行的内容
         */
        Row rowLast = sheet.createRow(rowIndex);
        rowLast.setHeightInPoints(13.5f);
        Cell rowLastcell0 = rowLast.createCell(0, CellType.STRING);
        rowLastcell0.setCellValue("上班\"√\"大夜班\"√\"加班\"√\"休息\"●\"事假\"×\"病假\"△\"旷工\"○\"迟到\"★\"早退\"▲\"漏打卡\"⊙\"婚嫁\"+\"丧假\"±\"离职\"＃\"工伤生育假\"※\"");
        // 设置最后一行标题的单元格格式
        CellStyle titleCellStyleLast = wb.createCellStyle();
        titleCellStyleLast.setAlignment(HorizontalAlignment.CENTER);
        titleCellStyleLast.setVerticalAlignment(VerticalAlignment.CENTER);
        Font fontLast = wb.createFont();
        fontLast.setFontName("宋体");
        fontLast.setFontHeightInPoints((short) 11);
        titleCellStyleLast.setFont(fontLast);
        rowLastcell0.setCellStyle(titleCellStyleLast);
        CellRangeAddress regionLast = new CellRangeAddress(rowIndex, rowIndex, 0, punchInRecords.size() + 15);
        sheet.addMergedRegion(regionLast);

        return wb;
    }

    /**
     * 生成x月打卡时间【表二】
     *
     * @param dingdingDailyStatistics 用户的打卡记录
     * @param wb                      输出的excel
     * @return 输出excel
     */
    private Workbook generatorXmonthClockInTime(List<DingdingDailyStatistics> dingdingDailyStatistics, Workbook wb) {
        if (null == wb || CollectionUtils.isEmpty(dingdingDailyStatistics)) {
            return wb;
        }

        // 创建工作表
        Sheet sheet = wb.getSheet(getMonthStr() + "月打卡时间");
        // 获取日期列表【就按照第一个员工的打卡记录来即可】
        List<DingdingPunchInRecord> punchInRecords = dingdingDailyStatistics.get(0).getPunchInRecords();
        // 行标【表格内容从第0行开始写】
        int rowIndex = 0;
        /**
         * *****************************************************************************************************
         * 开始填充数据
         * *****************************************************************************************************
         */

        /**
         * 1.设置第一行的标题
         */
        String title = clockInTimeAndDate.replace("每日统计表", "打卡时间表");
        Row row0 = sheet.createRow(rowIndex++);
        // 设置第一行的行高
        row0.setHeightInPoints(51.2f);
        // 设置第一行的内容
        Cell row0cell0 = row0.createCell(0, CellType.STRING);
        row0cell0.setCellValue(title);
        // 设置第一行标题的单元格格式
        CellStyle titleCellStyle = wb.createCellStyle();
        // 设置前景色
        titleCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        titleCellStyle.setFillForegroundColor(HSSFColor.LIGHT_TURQUOISE.index);
        // 设置字体
        Font font = wb.createFont();
        font.setFontName("新宋体");
        font.setFontHeightInPoints((short) 24);
        font.setBold(true);
        titleCellStyle.setFont(font);
        // 设置单元格边框
        titleCellStyle.setBorderBottom(BorderStyle.THIN);
        titleCellStyle.setBorderTop(BorderStyle.THIN);
        titleCellStyle.setBorderLeft(BorderStyle.THIN);
        titleCellStyle.setBorderRight(BorderStyle.THIN);
        // 垂直居中
        titleCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        row0cell0.setCellStyle(titleCellStyle);
        // 合并单元格【整个表格的列数登录当月的天数+4个固定的前置列】
        CellRangeAddress region = new CellRangeAddress(0, 0, 0, punchInRecords.size() + 3);
        sheet.addMergedRegion(region);

        /**
         * 2.设置第二行的标题
         */
        Row row1 = sheet.createRow(rowIndex++);
        // 设置行高
        row1.setHeightInPoints(51.2f);
        // 设置第二行标题的单元格格式
        CellStyle titleCellStyle2 = wb.createCellStyle();
        // 设置前景色
        titleCellStyle2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        titleCellStyle2.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
        // 水平居中
        titleCellStyle2.setAlignment(HorizontalAlignment.CENTER);
        // 垂直居中
        titleCellStyle2.setVerticalAlignment(VerticalAlignment.CENTER);
        // 表格边框
        titleCellStyle2.setBorderBottom(BorderStyle.THIN);
        titleCellStyle2.setBorderTop(BorderStyle.THIN);
        titleCellStyle2.setBorderLeft(BorderStyle.THIN);
        titleCellStyle2.setBorderRight(BorderStyle.THIN);
        // 字体
        Font font2 = wb.createFont();
        font2.setFontName("新宋体");
        font2.setFontHeightInPoints((short) 12);
        font2.setBold(true);
        titleCellStyle2.setFont(font2);
        // 设置第二行内容
        Cell row1cell0 = row1.createCell(0, CellType.STRING);
        // 第一列姓名
        row1cell0.setCellValue("姓名");
        row1cell0.setCellStyle(titleCellStyle2);
        Cell row1cell1 = row1.createCell(1, CellType.STRING);
        // 第二列考勤组
        row1cell1.setCellValue("考勤组");
        row1cell1.setCellStyle(titleCellStyle2);
        Cell row1cell2 = row1.createCell(2, CellType.STRING);
        // 第三列部门
        row1cell2.setCellValue("部门");
        row1cell2.setCellStyle(titleCellStyle2);
        Cell row1cell3 = row1.createCell(3, CellType.STRING);
        // 第四列职位
        row1cell3.setCellValue("职位");
        row1cell3.setCellStyle(titleCellStyle2);

        // 获取日期列表【日期对应的省略日期提示】
        Map<String, String> datatimeStrMap = punchInRecords.stream().map(p -> p.getDateTime()).collect(Collectors.toMap(
                p -> p.substring(0, p.indexOf(" ")),
                p -> {
                    if (p.indexOf("六") > -1 || p.indexOf("日") > -1) {
                        return p.substring(p.indexOf("星期") + 2);
                    }
                    p = p.substring(p.lastIndexOf("-") + 1, p.indexOf("星期"));
                    return p.charAt(0) == '0' ? p.substring(1) : p;
                },
                (l1, l2) -> l2));
        // 从第五列开始补充第二号的日期标题
        int i = 4;
        TreeSet<String> treeset = new TreeSet<>(datatimeStrMap.keySet());
        for (String key : treeset) {
            Cell row1celli = row1.createCell(i++, CellType.STRING);
            row1celli.setCellValue(datatimeStrMap.get(key));
            row1celli.setCellStyle(titleCellStyle2);
        }

        /**
         * 3. 开始填充真实数据
         */
        // 普通内容的单元格样式
        CellStyle contentCellStyle = wb.createCellStyle();
        contentCellStyle.setWrapText(true);
        contentCellStyle.setAlignment(HorizontalAlignment.CENTER);
        contentCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        Font font3 = wb.createFont();
        font3.setFontName("黑体");
        font3.setFontHeightInPoints((short) 12);
        contentCellStyle.setFont(font3);
        contentCellStyle.setBorderBottom(BorderStyle.THIN);
        contentCellStyle.setBorderTop(BorderStyle.THIN);
        contentCellStyle.setBorderLeft(BorderStyle.THIN);
        contentCellStyle.setBorderRight(BorderStyle.THIN);
        // 定义红色单元格格式
        CellStyle redContentCellStyle = wb.createCellStyle();
        redContentCellStyle.setWrapText(true);
        redContentCellStyle.setAlignment(HorizontalAlignment.LEFT);
        redContentCellStyle.setVerticalAlignment(VerticalAlignment.TOP);
        redContentCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        redContentCellStyle.setFillForegroundColor(HSSFColor.RED.index);
        Font font4 = wb.createFont();
        font4.setFontName("黑体");
        font4.setFontHeightInPoints((short) 12);
        redContentCellStyle.setFont(font4);
        redContentCellStyle.setBorderBottom(BorderStyle.THIN);
        redContentCellStyle.setBorderTop(BorderStyle.THIN);
        redContentCellStyle.setBorderLeft(BorderStyle.THIN);
        redContentCellStyle.setBorderRight(BorderStyle.THIN);
        // 定义无颜色单元格格式
        CellStyle noColorContentCellStyle = wb.createCellStyle();
        noColorContentCellStyle.setWrapText(true);
        noColorContentCellStyle.setAlignment(HorizontalAlignment.LEFT);
        noColorContentCellStyle.setVerticalAlignment(VerticalAlignment.TOP);
        noColorContentCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font5 = wb.createFont();
        font5.setFontName("黑体");
        font5.setFontHeightInPoints((short) 12);
        noColorContentCellStyle.setFont(font5);
        noColorContentCellStyle.setBorderBottom(BorderStyle.THIN);
        noColorContentCellStyle.setBorderTop(BorderStyle.THIN);
        noColorContentCellStyle.setBorderLeft(BorderStyle.THIN);
        noColorContentCellStyle.setBorderRight(BorderStyle.THIN);
        // 定义蓝色单元格格式
        CellStyle blueContentCellStyle = wb.createCellStyle();
        blueContentCellStyle.setWrapText(true);
        blueContentCellStyle.setAlignment(HorizontalAlignment.LEFT);
        blueContentCellStyle.setVerticalAlignment(VerticalAlignment.TOP);
        blueContentCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        blueContentCellStyle.setFillForegroundColor(HSSFColor.LIGHT_BLUE.index);
        Font font6 = wb.createFont();
        font6.setFontName("黑体");
        font6.setFontHeightInPoints((short) 12);
        blueContentCellStyle.setFont(font6);
        blueContentCellStyle.setBorderBottom(BorderStyle.THIN);
        blueContentCellStyle.setBorderTop(BorderStyle.THIN);
        blueContentCellStyle.setBorderLeft(BorderStyle.THIN);
        blueContentCellStyle.setBorderRight(BorderStyle.THIN);
        // 开始遍历员工打卡记录
        for (DingdingDailyStatistics data : dingdingDailyStatistics) {
            String userId = data.getUserId();
            // 加班次数
            int workOvertimeCount = 0;
            // 大夜班次数
            int nightShiftCount = 0;

            // 一个员工创建一行打卡记录数据
            Row rowF = sheet.createRow(rowIndex++);
            // 设置行高
            rowF.setHeightInPoints(51.2f);
            // 开始创建改行的每个单元格
            int cellIndex = 0;
            // 第一列姓名
            Cell rowFcell0 = rowF.createCell(cellIndex++, CellType.STRING);
            rowFcell0.setCellValue(data.getName());
            rowFcell0.setCellStyle(contentCellStyle);
            // 第二列考勤组
            Cell rowFcell1 = rowF.createCell(cellIndex++, CellType.STRING);
            rowFcell1.setCellValue(data.getAttendanceContent());
            rowFcell1.setCellStyle(contentCellStyle);
            // 第三列部门
            Cell rowFcell2 = rowF.createCell(cellIndex++, CellType.STRING);
            rowFcell2.setCellValue(data.getDeptName());
            rowFcell2.setCellStyle(contentCellStyle);
            // 第四列职位
            Cell rowFcell3 = rowF.createCell(cellIndex++, CellType.STRING);
            rowFcell3.setCellValue(data.getJobName());
            rowFcell3.setCellStyle(contentCellStyle);
            // 开始处理日期数据
            List<DingdingPunchInRecord> recordList = data.getPunchInRecords();
            if (CollectionUtils.isNotEmpty(recordList)) {
                for (DingdingPunchInRecord record : recordList) {
                    StringBuilder datatimeStr = new StringBuilder("");
                    if (StringUtils.isNotEmpty(record.getGoToWorkClockInTime1())) {
                        datatimeStr.append("\r\n").append(record.getGoToWorkClockInTime1()
                                .replace("昨日 ", "").replace("次日 ", ""));
                    }
                    if (StringUtils.isNotEmpty(record.getGoOffWorkClockInTime1())) {
                        datatimeStr.append("\r\n").append(record.getGoOffWorkClockInTime1()
                                .replace("昨日 ", "").replace("次日 ", ""));
                    }
                    if (StringUtils.isNotEmpty(record.getGoToWorkClockInTime2())) {
                        datatimeStr.append("\r\n").append(record.getGoToWorkClockInTime2()
                                .replace("昨日 ", "").replace("次日 ", ""));
                    }
                    if (StringUtils.isNotEmpty(record.getGoOffWorkClockInTime2())) {
                        datatimeStr.append("\r\n").append(record.getGoOffWorkClockInTime2()
                                .replace("昨日 ", "").replace("次日 ", ""));
                    }
                    Cell rowFcelli = rowF.createCell(cellIndex++, CellType.STRING);
                    rowFcelli.setCellValue(datatimeStr.toString().replaceFirst("\r\n", ""));

                    /**
                     * 判断表格颜色：
                     * 1. 蓝色：表示上夜班，打卡时间段在晚上12点到早上八点
                     * 2. 红色：表示加班，超过晚上八点打下班卡，表示加班
                     * 3. 无颜色：其他情况
                     * 4. 只有客服才有夜班
                     */
                    rowFcelli.setCellStyle(noColorContentCellStyle);
                    if (getOvertimeMarking(record.getGoOffWorkClockInTime1(), record.getGoToWorkClockInTime2(), record.getGoOffWorkClockInTime2())) {
                        rowFcelli.setCellStyle(redContentCellStyle);
                        workOvertimeCount++;
                    }
                    if (data.getAttendanceContent().equals("客服")) {
                        if (getNightShiftSign(record.getGoOffWorkClockInTime1(), record.getGoToWorkClockInTime2(), record.getGoOffWorkClockInTime2())) {
                            rowFcelli.setCellStyle(blueContentCellStyle);
                            nightShiftCount++;
                        }
                    }
                }
            }

            // 加班次数
            workOvertimeMap.put(userId, workOvertimeCount);
            // 大夜班次数
            nightShiftMap.put(userId, nightShiftCount);
        }

        return wb;
    }

    /**
     * 获取excel单元格的字符串的内容
     *
     * @param cell
     * @return
     */
    private String getCellStringValue(Cell cell) {
        String value = "";
        if (null == cell) {
            return value;
        }

        CellType cellTypeEnum = cell.getCellTypeEnum();
        if (cellTypeEnum == CellType.NUMERIC) {
            cell.setCellType(CellType.STRING);
        }
        return cell.getStringCellValue();
    }

    /**
     * @param
     * @return java.lang.String
     * @description 获取表格月份
     * @author cuiqiongyu
     * @date 22:06 2020-10-19
     **/
    private String getMonthStr() {
        String monthStr = "x";
        // 从clockInTimeAndDate解析x
        monthStr = clockInTimeAndDate.substring(clockInTimeAndDate.indexOf("：") + 1, clockInTimeAndDate.indexOf(" 至"));
        monthStr = monthStr.substring(monthStr.indexOf("-") + 1, monthStr.lastIndexOf("-"));
        monthStr = (monthStr.charAt(0) == '0') ? monthStr.substring(1) : monthStr;
        return monthStr;
    }

    /**
     * 获取加班标记 超过晚上八点打下班卡
     *
     * @param time2 打卡时间2
     * @param time3 打卡时间3
     * @param time4 打卡时间4
     * @return 是否加班
     */
    private boolean getOvertimeMarking(String time2, String time3, String time4) {
        try {
            time2 = time2.replace("昨日 ", "").replace("次日 ", "");
            time3 = time3.replace("昨日 ", "").replace("次日 ", "");
            time4 = time4.replace("昨日 ", "").replace("次日 ", "");

            for (String time : Lists.newArrayList(time4, time3, time2)) {
                if (StringUtils.isNotEmpty(time)) {
                    // 拆分时间
                    String[] split = time.split(":");
                    int num1 = Integer.valueOf(split[0]);
                    if (num1 == 20) {
                        return true;
                    }
                }
            }
            return false;
        } catch (Exception e) {
            return false;
        }
    }

    /**
     * 获取夜班标记 晚上00点前打上班卡，早上八点之后打下班卡
     *
     * @param time2 打卡时间2
     * @param time3 打卡时间3
     * @param time4 打卡时间4
     * @return 是否加班
     */
    private boolean getNightShiftSign(String time2, String time3, String time4) {
        try {
            time2 = time2.replace("昨日 ", "").replace("次日 ", "");
            time3 = time3.replace("昨日 ", "").replace("次日 ", "");
            time4 = time4.replace("昨日 ", "").replace("次日 ", "");

            List<String> list = Lists.newArrayList(time2, time3, time4);
            for (String time : list) {
                if (StringUtils.isNotEmpty(time)) {
                    // 拆分时间
                    String[] split = time.split(":");
                    int num1 = Integer.valueOf(split[0]);
                    if (num1 == 8) {
                        return true;
                    }
                }
            }
        } catch (Exception e) {
            return false;
        }

        return false;
    }

    /**
     * 获取下载路径
     *
     * @param filename 文件名称
     */
    public String getAbsoluteFile(String filename) {
        String downloadPath = RuoYiConfig.getDownloadPath() + filename;
        File desc = new File(downloadPath);
        if (!desc.getParentFile().exists()) {
            desc.getParentFile().mkdirs();
        }
        return downloadPath;
    }

}
