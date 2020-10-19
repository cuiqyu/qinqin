package com.littlebayo.project.qinq.pengda.service.impl;

import com.google.common.collect.Lists;
import com.littlebayo.common.exception.BusinessException;
import com.littlebayo.framework.web.domain.AjaxResult;
import com.littlebayo.project.qinq.pengda.domain.DingdingDailyStatistics;
import com.littlebayo.project.qinq.pengda.domain.DingdingPunchInRecord;
import com.littlebayo.project.qinq.pengda.service.AttendanceService;
import javafx.scene.paint.Color;
import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.swing.filechooser.FileSystemView;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
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
        OutputStream out = new FileOutputStream(desktopUrl + "/" + fileRealName + ATTENDANCE_REPORT_SUFFIX + suffix);
        wb.write(out);
        return AjaxResult.success("钉钉考勤文件上传成功，成功解析钉钉考勤文件<font color='blue'>[" + filename + "]</font>，<br/>" +
                "<font color='red'>解析后生成的考勤报表的路径为桌面路径，文件名为" + fileRealName + ATTENDANCE_REPORT_SUFFIX + suffix + "</font>。");
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
        String monthStr = "x";
        // 从clockInTimeAndDate解析x
        monthStr = clockInTimeAndDate.substring(clockInTimeAndDate.indexOf("：") + 1, clockInTimeAndDate.indexOf(" 至"));
        monthStr = monthStr.substring(monthStr.indexOf("-") + 1, monthStr.lastIndexOf("-"));
        monthStr = (monthStr.charAt(0) == '0') ? monthStr.substring(1) : monthStr;
        Sheet sheet = wb.createSheet(monthStr + "月考勤表统计");

        // TODO

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
        String monthStr = "x";
        // 从clockInTimeAndDate解析x
        monthStr = clockInTimeAndDate.substring(clockInTimeAndDate.indexOf("：") + 1, clockInTimeAndDate.indexOf(" 至"));
        monthStr = monthStr.substring(monthStr.indexOf("-") + 1, monthStr.lastIndexOf("-"));
        monthStr = (monthStr.charAt(0) == '0') ? monthStr.substring(1) : monthStr;
        Sheet sheet = wb.createSheet(monthStr + "月打卡时间");

        // 获取日期列表
        List<DingdingPunchInRecord> punchInRecords = dingdingDailyStatistics.get(0).getPunchInRecords();
        // 行标
        int rowIndex = 0;
        /**
         * 开始填充数据
         */
        // 1.设置第一行的标题
        String title = clockInTimeAndDate.replace("每日统计表", "打卡时间表");
        Row row = sheet.createRow(rowIndex++);
        row.createCell(0, CellType.STRING).setCellValue(title);
        // 设置第一行标题的单元格格式
        CellStyle titleCellStyle = wb.createCellStyle();
        titleCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        titleCellStyle.setFillForegroundColor(HSSFColor.LIGHT_TURQUOISE.index);
        Font font = wb.createFont();
        font.setFontName("新宋体");
        font.setFontHeightInPoints((short) 24);
        font.setBold(true);
        titleCellStyle.setFont(font);
        row.setRowStyle(titleCellStyle);
        CellRangeAddress region = new CellRangeAddress(0, 0, 0, punchInRecords.size() + 4);
        sheet.addMergedRegion(region);

        // 2. 设置第二行标题的单元格格式
        Row row1 = sheet.createRow(rowIndex++);
        CellStyle titleCellStyle2 = wb.createCellStyle();
        titleCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        titleCellStyle.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
        Font font2 = wb.createFont();
        font2.setFontName("新宋体");
        font2.setFontHeightInPoints((short) 12);
        font2.setBold(true);
        titleCellStyle2.setFont(font);
        row1.setRowStyle(titleCellStyle2);
        // 设置第二行内容
        row1.createCell(0, CellType.STRING).setCellValue("姓名");
        row1.createCell(1, CellType.STRING).setCellValue("考勤组");
        row1.createCell(2, CellType.STRING).setCellValue("部门");
        row1.createCell(3, CellType.STRING).setCellValue("职位");
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
            row1.createCell(i++, CellType.STRING).setCellValue(datatimeStrMap.get(key));
        }

        // 3. 开始填充内容数据
        for (DingdingDailyStatistics dingdingDailyStatistic : dingdingDailyStatistics) {
            Row rowF = sheet.createRow(rowIndex++);
            int cellIndex = 0;
            rowF.createCell(cellIndex++, CellType.STRING).setCellValue(dingdingDailyStatistic.getName());
            rowF.createCell(cellIndex++, CellType.STRING).setCellValue(dingdingDailyStatistic.getAttendanceContent());
            rowF.createCell(cellIndex++, CellType.STRING).setCellValue(dingdingDailyStatistic.getDeptName());
            rowF.createCell(cellIndex++, CellType.STRING).setCellValue(dingdingDailyStatistic.getJobName());
            // 开始处理日期

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

}
