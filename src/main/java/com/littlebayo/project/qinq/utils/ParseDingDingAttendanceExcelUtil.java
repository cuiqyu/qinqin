package com.littlebayo.project.qinq.utils;

import com.littlebayo.common.exception.BusinessException;
import com.littlebayo.project.qinq.pengda.domain.DingdingMonthlyStatistics;
import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 解析钉钉考情报表工具
 *
 * @author cuiqiongyu
 * @date 2020/10/29 10:58
 */
public class ParseDingDingAttendanceExcelUtil {

    private final static Logger logger = LoggerFactory.getLogger(ParseDingDingAttendanceExcelUtil.class);

    /**
     * 解析钉钉考勤报表每月统计（第一张表格）
     *
     * @param inputStream 文件输入流
     * @return 解析结果
     * @throws IOException
     * @throws InvalidFormatException
     */
    public static List<DingdingMonthlyStatistics> parseDingdingExcelMonthlyStatisticsData(InputStream inputStream, String sheetName)
            throws IOException, InvalidFormatException {
        // 工作簿对象
        Workbook wb = WorkbookFactory.create(inputStream);
        // 获取【月度汇总】表格
        Sheet sheet = wb.getSheet(sheetName);
        if (null == sheet) {
            logger.error("解析钉钉考勤报表失败，未找到报表文件中的[月度汇总]表格");
            throw new BusinessException("未找到报表文件中的[月度汇总]表格");
        }

        // 月度汇总数据内容
        List<DingdingMonthlyStatistics> dataList = new ArrayList<>();
        // 获取总行数【表头占用了四行】
        int rows = sheet.getPhysicalNumberOfRows();
        if (rows <= 4) {
            return dataList;
        }
        // 开始解析第三行的标题
        int threeCellIndex = 0;
        String currentCellValue = "";
        Map<String, Integer> titleMap = new HashMap<>();
        Row threeRow = sheet.getRow(2);
        while (threeCellIndex < 80 && !currentCellValue.equals("请假")) {
            Cell cell = threeRow.getCell(threeCellIndex);
            currentCellValue = cell.getStringCellValue();
            titleMap.put(currentCellValue, threeCellIndex);
            threeCellIndex++;
        }
        // 找到请假的合并单元格的位置
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        if (CollectionUtils.isNotEmpty(mergedRegions)) {
            CellRangeAddress cellRangeAddress =
                    mergedRegions.stream().filter(m -> m.getFirstRow() == 2 && m.getLastRow() == 2 && m.getFirstColumn() == titleMap.get("请假")).findFirst().orElse(null);
            if (null != cellRangeAddress) {
                Integer firstColumnIndex = titleMap.get("请假");
                int lastColumnIndex = cellRangeAddress.getLastColumn();
                titleMap.remove("请假");
                Row fourRow = sheet.getRow(3);
                for (int fourCellIndex = firstColumnIndex; fourCellIndex <= lastColumnIndex; fourCellIndex++) {
                    Cell cell = fourRow.getCell(fourCellIndex);
                    titleMap.put(cell.getStringCellValue(), fourCellIndex);
                }
            }
        }
        // 获取titleMap的size，表示每次打卡记录从那一列开始
        int everydaySummaryCellIndex = titleMap.size();

        // 开始遍历数据，从第五行【行标从第4】开始
        for (int i = 4; i < rows; i++) {
            Row rowi = sheet.getRow(i);
            Map<String, String> cellValueMap = new HashMap<>();
            for (Map.Entry<String, Integer> entry : titleMap.entrySet()) {
                cellValueMap.put(entry.getKey(), getCellStringValue(rowi.getCell(entry.getValue())));
            }
            DingdingMonthlyStatistics monthlyStatistics = new DingdingMonthlyStatistics();
            monthlyStatistics.setName(cellValueMap.get("姓名"));
            monthlyStatistics.setAttendanceContent(cellValueMap.get("考勤组"));
            monthlyStatistics.setDeptName(cellValueMap.get("部门"));
            monthlyStatistics.setJobNumber(cellValueMap.get("工号"));
            monthlyStatistics.setUserId(cellValueMap.get("UserId"));
            monthlyStatistics.setChuqinTianshu(cellValueMap.get("出勤天数"));
            monthlyStatistics.setXiuxiTianshu(cellValueMap.get("休息天数"));
            monthlyStatistics.setGongzuoShichang(cellValueMap.get("工作时长"));
            monthlyStatistics.setChidaoCishu(cellValueMap.get("迟到次数"));
            monthlyStatistics.setChidaoShichang(cellValueMap.get("迟到时长"));
            monthlyStatistics.setYanzhongChidaoCishu(cellValueMap.get("严重迟到次数"));
            monthlyStatistics.setYanzhongChidaoShichang(cellValueMap.get("严重迟到时长"));
            monthlyStatistics.setKuanggongChidaoTianshu(cellValueMap.get("旷工迟到天数"));
            monthlyStatistics.setZaotuiCishu(cellValueMap.get("早退次数"));
            monthlyStatistics.setZaotuiShichang(cellValueMap.get("早退时长"));
            monthlyStatistics.setShangbanQuekaCishu(cellValueMap.get("上班缺卡次数"));
            monthlyStatistics.setXiabanQuekaCishu(cellValueMap.get("下班缺卡次数"));
            monthlyStatistics.setKuangggongTianshu(cellValueMap.get("旷工天数"));
            monthlyStatistics.setWaichu(cellValueMap.get("外出"));
            monthlyStatistics.setSangjia(cellValueMap.get("丧假(天)"));
            monthlyStatistics.setShijia(cellValueMap.get("事假"));
            monthlyStatistics.setBingjia(cellValueMap.get("病假"));
            monthlyStatistics.setTiaoxiu(cellValueMap.get("调休"));

            // 每日打卡总结
            HashMap<Integer, String> hashMap = new HashMap<>();
            // 假设每个月按照最大的天数 31天来算
            for (int m = 0; m < 31; m++) {
                hashMap.put(m + 1, getCellStringValue(rowi.getCell(m + everydaySummaryCellIndex)));
            }
            monthlyStatistics.setEverydaySummary(hashMap);

            dataList.add(monthlyStatistics);
        }

        return dataList;
    }

    /**
     * 获取excel单元格的字符串的内容
     *
     * @param cell
     * @return
     */
    public static String getCellStringValue(Cell cell) {
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
