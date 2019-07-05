package com.example.code.excel;

import cn.hutool.json.JSONUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author Lv Jie
 * @version 1.0.0
 * @desc TODO
 * @create 2019/6/28 15:22
 */

@Slf4j
public class ImportExcelUtil {
	private final static String excel2003L = ".xls";
	private final static String excel2007U = ".xlsx";

	public static List<Map<String, Object>> parseExcel(String fileName) throws Exception {
		// 根据文件名来创建Excel工作薄
		Workbook work = getWorkbook(fileName);
		if (null == work) {
			throw new Exception("创建Excel工作薄为空！");
		}
		Sheet sheet = null;
		Row row = null;
		Cell cell = null;
		// 返回数据
		List<Map<String, Object>> ls = new ArrayList<Map<String, Object>>();

		// 遍历Excel中所有的sheet
		for (int i = 0; i < work.getNumberOfSheets(); i++) {
			sheet = work.getSheetAt(i);
			if (sheet == null){
				continue;
			}
			// 取第一行标题
			row = sheet.getRow(0);
			String title[] = null;
			if (row != null) {
				title = new String[row.getLastCellNum()];

				for (int y = row.getFirstCellNum(); y < row.getLastCellNum(); y++) {
					cell = row.getCell(y);
					title[y] = (String) getCellValue(cell);
				}

			} else {
				continue;
			}
			log.info("{}", JSONUtil.toJsonStr(title));

			// 遍历当前sheet中的所有行
			for (int j = 1; j < sheet.getLastRowNum() + 1; j++) {
				row = sheet.getRow(j);
				if (row == null) {
					continue;
				}
				Map<String, Object> m = new HashMap<String, Object>();
				// 遍历所有的列
				for (int y = row.getFirstCellNum(); y < row.getLastCellNum(); y++) {
					cell = row.getCell(y);
					String key = title[y];
					// log.info(JSON.toJSONString(key));
					m.put(key, getCellValue(cell));
				}
				ls.add(m);
			}
		}
		work.close();
		return ls;
	}

	/**
	 * 描述：根据文件后缀，自适应上传文件的版本
	 * <p>
	 * ,fileName
	 *
	 * @return
	 * @throws Exception
	 */
	public static Workbook getWorkbook(String fileName) throws Exception {
		InputStream in = new FileInputStream(new File(fileName));
		Workbook wb = null;
		String fileType = fileName.substring(fileName.lastIndexOf("."));
		if (excel2003L.equals(fileType)) {
			wb = new HSSFWorkbook(in); // 2003-
		} else if (excel2007U.equals(fileType)) {
			wb = new XSSFWorkbook(in); // 2007+
		} else {
			throw new Exception("解析的文件格式有误！");
		}
		return wb;
	}

	/**
	 * 描述：对表格中数值进行格式化
	 *
	 * @param cell
	 * @return
	 */
	public static Object getCellValue(Cell cell) {
		Object value = null;
		DecimalFormat df = new DecimalFormat("0"); // 格式化number String字符
		SimpleDateFormat sdf = new SimpleDateFormat("yyy-MM-dd"); // 日期格式化
		DecimalFormat df2 = new DecimalFormat("0"); // 格式化数字

		switch (cell.getCellType()) {
			case Cell.CELL_TYPE_STRING:
				value = cell.getRichStringCellValue().getString();
				break;
			case Cell.CELL_TYPE_NUMERIC:
				if ("General".equals(cell.getCellStyle().getDataFormatString())) {
					value = df.format(cell.getNumericCellValue());
				} else if ("m/d/yy".equals(cell.getCellStyle().getDataFormatString())) {
					value = sdf.format(cell.getDateCellValue());
				} else {
					value = df2.format(cell.getNumericCellValue());
				}
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				value = cell.getBooleanCellValue();
				break;
			case Cell.CELL_TYPE_BLANK:
				value = "";
				break;
			default:
				break;
		}
		return value;
	}

}

