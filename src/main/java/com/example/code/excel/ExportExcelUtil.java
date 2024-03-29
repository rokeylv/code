package com.example.code.excel;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.OutputStream;
import java.util.List;
import java.util.Map;

/**
 * @author Lv Jie
 * @version 1.0.0
 * @desc TODO
 * @create 2019/7/3 15:58
 */
@Slf4j
public class ExportExcelUtil {

	/**
	 * 导出列表数据
	 * @param title 表头集合
	 * @param position  表头字段位置集合
	 * @param data 需要导出的数据
	 * 导出数据后在excel表格中左下角显示的工作簿名称（注意：不是导出后的文件名）
	 * @param outputStream 从controller层通过response获取到的输出流
	 */
	public static void exportDataToExcel(Map<String, String> title, Map<String, Integer> position, List<Map<String, Object>> data, OutputStream outputStream){
		if (data == null || data.size() < 1) {
			return;
		}
		try (Workbook workbook = new XSSFWorkbook()) {
			Sheet sheet = workbook.createSheet();
			Row header = sheet.createRow(0);
			// 设置表头样式
			CellStyle headerStyle = workbook.createCellStyle();
			headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
			headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			// 字体样式
			XSSFFont font = ((XSSFWorkbook) workbook).createFont();
			font.setFontName("Arial");
			font.setFontHeightInPoints((short)14);
			headerStyle.setFont(font);
			int col = 0;
			 // 遍历表头map集合
			for (String key: title.keySet()) {
				sheet.setColumnWidth(col, 6000);
				// 设置表格头部
				Cell headerCell = header.createCell(position.get(key));
				headerCell.setCellValue(title.get(key) + "");
				headerCell.setCellStyle(headerStyle);
				col++;
			}
			CellStyle style = workbook.createCellStyle();
			style.setWrapText(true);
			/*
			 * 遍历要导出列表的数据data 并与title的key相比较， 确认后插入值
			 * 创建列时，根据title的key然后将值插入到对应的列中（position，dataMap，title三个集合的key值是一一对应的）
			 */
			if (data != null && data.size() > 0) {
				int r = 0;
				for (Map<String, Object> dataMap : data) {
					Row row = sheet.createRow(r + 1);
					for (String dkey : dataMap.keySet()) {
						for (String key : title.keySet()) {
							if (key.equals(dkey)) {
								Cell cell = row.createCell(position.get(key));
								cell.setCellValue(dataMap.get(dkey) + "");
								cell.setCellStyle(style);
								break;
							}
						}
					}
					r++;
				}
			}
			workbook.write(outputStream);
		} catch (Exception ex) {
			log.error("export data", ex);
			throw new RuntimeException("导出列表失败。");
		}
	}
}
