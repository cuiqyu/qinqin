package com.littlebayo.project.qinq.pengda.service.impl;

import com.google.common.collect.Lists;
import com.littlebayo.common.exception.BusinessException;
import com.littlebayo.common.utils.StringUtils;
import com.littlebayo.framework.config.RuoYiConfig;
import com.littlebayo.framework.web.domain.AjaxResult;
import com.littlebayo.project.qinq.pengda.domain.DingdingDailyStatistics;
import com.littlebayo.project.qinq.pengda.domain.DingdingPunchInRecord;
import com.littlebayo.project.qinq.pengda.service.AttendanceService;
import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.swing.filechooser.FileSystemView;
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
    /**
     * 生成的考勤报表添加的文件名后缀
     */
    private static final String ATTENDANCE_REPORT_SUFFIX = "_qinq";

    private static String clockInTimeAndDate = "每日统计表 统计日期：2020-09-01 至 2020-09-30";

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
        // 开始解析考勤表格文件，只解析第二个表格
        List<DingdingDailyStatistics> dingdingDailyStatistics = parseDingdingExcelDailyStatisticsData(file.getInputStream());
        // 获取桌面路径
        FileSystemView fsv = FileSystemView.getFileSystemView();
        String desktopUrl = fsv.getHomeDirectory().getPath();

        // -------开始分析并导出-------
        // 1. 创建工作簿对象
        Workbook wb = new SXSSFWorkbook(500);

        // 2. 生成x月考勤表统计【表一】
        wb = generatorXmonthCheckWorkAttendance(dingdingDailyStatistics, wb);

        // 3. 生成x月打卡时间【表二】
        wb = generatorXmonthClockInTime(dingdingDailyStatistics, wb);

        // 4. 输出工作表
        String fileName = fileRealName + ATTENDANCE_REPORT_SUFFIX + suffix;
        OutputStream out = new FileOutputStream(getAbsoluteFile(fileName));
        wb.write(out);
        return AjaxResult.success(fileName);
    }

    /**
     * 解析钉钉考勤报表每日统计（第二张表格）
     *
     * @param inputStream 文件输入流
     */
    private List<DingdingDailyStatistics> parseDingdingExcelDailyStatisticsData(InputStream inputStream) throws IOException, InvalidFormatException {
        // 工作簿对象
        Workbook wb = WorkbookFactory.create(inputStream);
        // 获取每日统计表格
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
     * 生成x月考勤表统计【表一】
     *
     * @param dingdingDailyStatistics 用户的打卡记录
     * @param wb                      输出的excel
     * @return 输出excel
     */
    private Workbook generatorXmonthCheckWorkAttendance(List<DingdingDailyStatistics> dingdingDailyStatistics, Workbook wb) {
        if (null == wb || CollectionUtils.isEmpty(dingdingDailyStatistics)) {
            return wb;
        }
        Sheet sheet = wb.createSheet(getMonthStr() + "月考勤表统计");
        // 设置列宽
        sheet.setColumnWidth(0, 5 * 256);
        sheet.setColumnWidth(1, 14 * 256);
        sheet.setColumnWidth(2, 5 * 256);
        sheet.setColumnWidth(3, 8 * 256);
        sheet.setColumnWidth(4, 6 * 256);
        sheet.setColumnWidth(5, 6 * 256);
        sheet.setColumnWidth(6, 6 * 256);
        sheet.setColumnWidth(7, 3 * 256);
        sheet.setColumnWidth(8, 3 * 256);
        sheet.setColumnWidth(9, 3 * 256);
        sheet.setColumnWidth(10, 4 * 256);
        sheet.setColumnWidth(11, 3 * 256);
        sheet.setColumnWidth(12, 3 * 256);
        sheet.setColumnWidth(13, 3 * 256);
        sheet.setColumnWidth(14, 3 * 256);
        sheet.setColumnWidth(15, 5 * 256);
        for (int i = 0; i < 31; i++) {
            sheet.setColumnWidth(16 + i, 3 * 256);
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
        Cell row2Cell0 = row2.createCell(row2cellIndex++, CellType.STRING);
        row2Cell0.setCellValue("序号");
        row2Cell0.setCellStyle(titleCellStyle2);
        Cell row2Cell1 = row2.createCell(row2cellIndex++, CellType.STRING);
        row2Cell1.setCellValue("姓名");
        row2Cell1.setCellStyle(titleCellStyle2);
        Cell row2Cell2 = row2.createCell(row2cellIndex++, CellType.STRING);
        row2Cell2.setCellValue("应出勤天数");
        row2Cell2.setCellStyle(titleCellStyle2);
        Cell row2Cell3 = row2.createCell(row2cellIndex++, CellType.STRING);
        row2Cell3.setCellValue("实际出勤天数");
        row2Cell3.setCellStyle(titleCellStyle2);
        Cell row2Cell4 = row2.createCell(row2cellIndex++, CellType.STRING);
        row2Cell4.setCellValue("缺勤");
        row2Cell4.setCellStyle(titleCellStyle2);
        Cell row2Cell5 = row2.createCell(row2cellIndex++, CellType.STRING);
        row2Cell5.setCellValue("平时加班天数");
        row2Cell5.setCellStyle(titleCellStyle2);
        Cell row2Cell6 = row2.createCell(row2cellIndex++, CellType.STRING);
        row2Cell6.setCellValue("事假天数");
        row2Cell6.setCellStyle(titleCellStyle2);
        Cell row2Cell7 = row2.createCell(row2cellIndex++, CellType.STRING);
        row2Cell7.setCellValue("产假天数");
        row2Cell7.setCellStyle(titleCellStyle2);
        Cell row2Cell8 = row2.createCell(row2cellIndex++, CellType.STRING);
        row2Cell8.setCellValue("丧假天数");
        row2Cell8.setCellStyle(titleCellStyle2);
        Cell row2Cell9 = row2.createCell(row2cellIndex++, CellType.STRING);
        row2Cell9.setCellValue("婚假天数");
        row2Cell9.setCellStyle(titleCellStyle2);
        Cell row2Cell10 = row2.createCell(row2cellIndex++, CellType.STRING);
        row2Cell10.setCellValue("病假天数");
        row2Cell10.setCellStyle(titleCellStyle2);
        Cell row2Cell11 = row2.createCell(row2cellIndex++, CellType.STRING);
        row2Cell11.setCellValue("迟到早退次数");
        row2Cell11.setCellStyle(titleCellStyle2);
        Cell row2Cell12 = row2.createCell(row2cellIndex++, CellType.STRING);
        row2Cell12.setCellValue("漏打卡次数");
        row2Cell12.setCellStyle(titleCellStyle2);
        Cell row2Cell13 = row2.createCell(row2cellIndex++, CellType.STRING);
        row2Cell13.setCellValue("大夜班补贴");
        row2Cell13.setCellStyle(titleCellStyle2);
        Cell row2Cell14 = row2.createCell(row2cellIndex++, CellType.STRING);
        row2Cell14.setCellValue("加班补贴");
        row2Cell14.setCellStyle(titleCellStyle2);
        Cell row2Cell15 = row2.createCell(row2cellIndex++, CellType.STRING);
        row2Cell15.setCellValue("日期时间");
        row2Cell15.setCellStyle(titleCellStyle2);
        List<String> list = datatimeStrMap.keySet().stream().sorted(Comparator.comparing(Integer::valueOf)).collect(Collectors.toList());
        for (String key : list) {
            row2.createCell(row2cellIndex++, CellType.STRING).setCellValue(key);
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

        // 开始填充数据
        for (int i = 0; i < dingdingDailyStatistics.size(); i++) {
            Row rown = sheet.createRow(rowIndex++);
            int rowncellIndex = 0;
            DingdingDailyStatistics statistics = dingdingDailyStatistics.get(i);
            Cell rownCell1 = rown.createCell(rowncellIndex++, CellType.STRING);
            rownCell1.setCellValue(i + 1);
            rownCell1.setCellStyle(titleCellStyle2);
            Cell rownCell2 = rown.createCell(rowncellIndex++, CellType.STRING);
            rownCell2.setCellValue(statistics.getName());
            rownCell2.setCellStyle(titleCellStyle2);
            // TODO
        }

        /**
         * 设置最后一行的内容
         */
        Row rowLast = sheet.createRow(rowIndex);
        rowLast.setHeightInPoints(13.5f);
        Cell rowLastcell0 = rowLast.createCell(0, CellType.STRING);
        rowLastcell0.setCellValue("上班\"√\"大夜班\"√\"加班\"√\"休息\"●\"事假\"×\"病假\"△\"旷工\"○\"迟到\"★\"早退\"▲\"漏打卡\"⊙\"婚嫁\"+\"丧假\"±\"离职\"＃\"工伤生育假\"※\"");
        // 设置第一行标题的单元格格式
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

        Sheet sheet = wb.createSheet(getMonthStr() + "月打卡时间");

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
        String title = clockInTimeAndDate.replace("每日统计表", "打卡时间表");
        Row row0 = sheet.createRow(rowIndex++);
        row0.setHeightInPoints(51.2f);
        Cell row0cell0 = row0.createCell(0, CellType.STRING);
        row0cell0.setCellValue(title);
        // 设置第一行标题的单元格格式
        CellStyle titleCellStyle = wb.createCellStyle();
        titleCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        titleCellStyle.setFillForegroundColor(HSSFColor.LIGHT_TURQUOISE.index);
        Font font = wb.createFont();
        font.setFontName("新宋体");
        font.setFontHeightInPoints((short) 24);
        font.setBold(true);
        titleCellStyle.setFont(font);
        titleCellStyle.setBorderBottom(BorderStyle.THIN);
        titleCellStyle.setBorderTop(BorderStyle.THIN);
        titleCellStyle.setBorderLeft(BorderStyle.THIN);
        titleCellStyle.setBorderRight(BorderStyle.THIN);
        titleCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        row0cell0.setCellStyle(titleCellStyle);
        CellRangeAddress region = new CellRangeAddress(0, 0, 0, punchInRecords.size() + 3);
        sheet.addMergedRegion(region);

        /**
         * 2.设置第二行的标题
         */
        // 设置第二行标题的单元格格式
        Row row1 = sheet.createRow(rowIndex++);
        row1.setHeightInPoints(51.2f);
        CellStyle titleCellStyle2 = wb.createCellStyle();
        titleCellStyle2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        titleCellStyle2.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
        titleCellStyle2.setAlignment(HorizontalAlignment.CENTER);
        titleCellStyle2.setBorderBottom(BorderStyle.THIN);
        titleCellStyle2.setBorderTop(BorderStyle.THIN);
        titleCellStyle2.setBorderLeft(BorderStyle.THIN);
        titleCellStyle2.setBorderRight(BorderStyle.THIN);
        Font font2 = wb.createFont();
        font2.setFontName("新宋体");
        font2.setFontHeightInPoints((short) 12);
        font2.setBold(true);
        titleCellStyle2.setFont(font2);
        titleCellStyle2.setVerticalAlignment(VerticalAlignment.CENTER);
        // 设置第二行内容
        Cell row1cell0 = row1.createCell(0, CellType.STRING);
        row1cell0.setCellValue("姓名");
        row1cell0.setCellStyle(titleCellStyle2);
        Cell row1cell1 = row1.createCell(1, CellType.STRING);
        row1cell1.setCellValue("考勤组");
        row1cell1.setCellStyle(titleCellStyle2);
        Cell row1cell2 = row1.createCell(2, CellType.STRING);
        row1cell2.setCellValue("部门");
        row1cell2.setCellStyle(titleCellStyle2);
        Cell row1cell3 = row1.createCell(3, CellType.STRING);
        row1cell3.setCellValue("职位");
        row1cell3.setCellStyle(titleCellStyle2);

        // 获取日期列表
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
        int i = 4;
        TreeSet<String> treeset = new TreeSet<>(datatimeStrMap.keySet());
        for (String key : treeset) {
            Cell row1celli = row1.createCell(i++, CellType.STRING);
            row1celli.setCellValue(datatimeStrMap.get(key));
            row1celli.setCellStyle(titleCellStyle2);
        }

        // 3. 开始填充内容数据
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

        for (DingdingDailyStatistics data : dingdingDailyStatistics) {
            Row rowF = sheet.createRow(rowIndex++);
            rowF.setHeightInPoints(51.2f);
            int cellIndex = 0;
            Cell rowFcell0 = rowF.createCell(cellIndex++, CellType.STRING);
            rowFcell0.setCellValue(data.getName());
            rowFcell0.setCellStyle(contentCellStyle);
            Cell rowFcell1 = rowF.createCell(cellIndex++, CellType.STRING);
            rowFcell1.setCellValue(data.getAttendanceContent());
            rowFcell1.setCellStyle(contentCellStyle);
            Cell rowFcell2 = rowF.createCell(cellIndex++, CellType.STRING);
            rowFcell2.setCellValue(data.getDeptName());
            rowFcell2.setCellStyle(contentCellStyle);
            Cell rowFcell3 = rowF.createCell(cellIndex++, CellType.STRING);
            rowFcell3.setCellValue(data.getJobName());
            rowFcell3.setCellStyle(contentCellStyle);
            // 开始处理日期
            List<DingdingPunchInRecord> recordList = data.getPunchInRecords();
            if (CollectionUtils.isNotEmpty(recordList)) {
                for (DingdingPunchInRecord record : recordList) {
                    StringBuilder datatimeStr = new StringBuilder("");
                    if (StringUtils.isNotEmpty(record.getGoToWorkClockInTime1())) {
                        datatimeStr.append("\r\n").append(record.getGoToWorkClockInTime1().replace("昨日 ", "").replace("次日 ", ""));
                    }
                    if (StringUtils.isNotEmpty(record.getGoOffWorkClockInTime1())) {
                        datatimeStr.append("\r\n").append(record.getGoOffWorkClockInTime1().replace("昨日 ", "").replace("次日 ", ""));
                    }
                    if (StringUtils.isNotEmpty(record.getGoToWorkClockInTime2())) {
                        datatimeStr.append("\r\n").append(record.getGoToWorkClockInTime2().replace("昨日 ", "").replace("次日 ", ""));
                    }
                    if (StringUtils.isNotEmpty(record.getGoOffWorkClockInTime2())) {
                        datatimeStr.append("\r\n").append(record.getGoOffWorkClockInTime2().replace("昨日 ", "").replace("次日 ", ""));
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
                    if (getOvertimeMarking(record.getGoToWorkClockInTime1(), record.getGoOffWorkClockInTime1(), record.getGoToWorkClockInTime2(), record.getGoOffWorkClockInTime2())) {
                        rowFcelli.setCellStyle(redContentCellStyle);
                    }
                    if (data.getAttendanceContent().equals("客服")) {
                        if (getNightShiftSign(record.getGoToWorkClockInTime1(), record.getGoOffWorkClockInTime1(), record.getGoToWorkClockInTime2(), record.getGoOffWorkClockInTime2())) {
                            rowFcelli.setCellStyle(blueContentCellStyle);
                        }
                    }
                }
            }
        }

        // TODO
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
     * @param time1 打卡时间1
     * @param time2 打卡时间2
     * @param time3 打卡时间3
     * @param time4 打卡时间4
     * @return 是否加班
     */
    private boolean getOvertimeMarking(String time1, String time2, String time3, String time4) {
        try {
            time1 = time1.replace("昨日 ", "").replace("次日 ", "");
            time2 = time2.replace("昨日 ", "").replace("次日 ", "");
            time3 = time3.replace("昨日 ", "").replace("次日 ", "");
            time4 = time4.replace("昨日 ", "").replace("次日 ", "");

            if (StringUtils.isNotEmpty(time4)) {
                // 拆分时间
                String[] split = time4.split(":");
                int num1 = Integer.valueOf(split[0]);
                if (num1 == 20) {
                    return true;
                }
            }
            if (StringUtils.isNotEmpty(time3)) {
                // 拆分时间
                String[] split = time3.split(":");
                int num1 = Integer.valueOf(split[0]);
                if (num1 == 20) {
                    return true;
                }
            }
            if (StringUtils.isNotEmpty(time2)) {
                // 拆分时间
                String[] split = time2.split(":");
                int num1 = Integer.valueOf(split[0]);
                if (num1 == 20) {
                    return true;
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
     * @param time1 打卡时间1
     * @param time2 打卡时间2
     * @param time3 打卡时间3
     * @param time4 打卡时间4
     * @return 是否加班
     */
    private boolean getNightShiftSign(String time1, String time2, String time3, String time4) {
        try {
            time1 = time1.replace("昨日 ", "").replace("次日 ", "");
            time2 = time2.replace("昨日 ", "").replace("次日 ", "");
            time3 = time3.replace("昨日 ", "").replace("次日 ", "");
            time4 = time4.replace("昨日 ", "").replace("次日 ", "");

/*            // 判断上班时间
            if (StringUtils.isNotEmpty(time1)) {
                // 拆分时间
                String[] split = time1.split(":");
                int num1 = Integer.valueOf(split[0]);
                if (num1 == 23 || num1 == 0) {
                    return true;
                }
            }*/
            if (StringUtils.isNotEmpty(time2)) {
                // 拆分时间
                String[] split = time2.split(":");
                int num1 = Integer.valueOf(split[0]);
                if (num1 == 8) {
                    return true;
                }
            }
            if (StringUtils.isNotEmpty(time3)) {
                // 拆分时间
                String[] split = time3.split(":");
                int num1 = Integer.valueOf(split[0]);
                if (num1 == 8) {
                    return true;
                }
            }
            // 判断下班时间
            if (StringUtils.isNotEmpty(time4)) {
                // 拆分时间
                String[] split = time4.split(":");
                int num1 = Integer.valueOf(split[0]);
                if (num1 == 8) {
                    return true;
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
    public String getAbsoluteFile(String filename)
    {
        String downloadPath = RuoYiConfig.getDownloadPath() + filename;
        File desc = new File(downloadPath);
        if (!desc.getParentFile().exists())
        {
            desc.getParentFile().mkdirs();
        }
        return downloadPath;
    }

}
