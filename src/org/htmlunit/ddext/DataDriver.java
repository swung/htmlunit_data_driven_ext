package org.htmlunit.ddext;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.gargoylesoftware.htmlunit.util.NameValuePair;

public class DataDriver
{

	private transient Collection<Object[]> data = null;
	private Class testCaseClass = null;

	public DataDriver(final InputStream excelInputStream, final Class testCaseClass) throws IOException
	{
		this.testCaseClass = testCaseClass;
		this.data = loadFromSpreadsheet(excelInputStream);
	}

	public Collection<Object[]> getData()
	{
		return data;
	}

	private Collection<Object[]> loadFromSpreadsheet(final InputStream excelFile) throws IOException
	{

		if (testCaseClass == null) {
			return null;
		}

		List<Object[]> result = new ArrayList<Object[]>();

		HSSFWorkbook workbook = new HSSFWorkbook(excelFile);

		List<String> sheetNames = new ArrayList<String>();
		for (Field f : testCaseClass.getDeclaredFields()) {
			if (f.getAnnotation(Sheet.class) != null) {
				sheetNames.add(f.getName());
			}
		}

		String[] sheetArray = sheetNames.toArray(new String[sheetNames.size()]);

		Map<String, Map<String, List<NameValuePair>>> listMap = new HashMap<String, Map<String, List<NameValuePair>>>();

		for (String sn : sheetNames) {
			HSSFSheet sheet = workbook.getSheet(sn);
			List<String> params = new ArrayList<String>();
			Map<String, List<NameValuePair>> mapListNVP = new HashMap<String, List<NameValuePair>>();

			int numberOfColumns = countNonEmptyColumns(sheet);

			for (Row row : sheet) {
				List<NameValuePair> listNVP = new ArrayList<NameValuePair>();
				if (row.getRowNum() == 0) {
					// Ignore row1 column1
					for (int column = 1; column < numberOfColumns; column++) {
						params.add(row.getCell(column).getStringCellValue().trim());
					}
				} else {
					if (isCommentOutRow(row) || isEmptyCell(row.getCell(0))) {
						continue;
					}

					// colum1 as ID
					String k = cellToString(row.getCell(0));
					int i = 1;
					for (String p : params) {
						Cell cell = row.getCell(i);
						if (isEmptyCell(cell)) {
							listNVP.add(new NameValuePair(p, ""));
						} else {
							listNVP.add(new NameValuePair(p, cellToString(cell)));
						}
						i++;
					}
					mapListNVP.put(k, listNVP);
				}
			}

			listMap.put(sn, mapListNVP);
		}

		Set<String> keys = listMap.get(sheetArray[0]).keySet();
		int size = listMap.keySet().size();
		Object[] item = new Object[size];
		for (String k : keys) {
			for (int j = 0; j < size; j++) {
				Object o = listMap.get(sheetArray[j]).get(k);
				if (o != null) {
					item[j] = o;
				} else {
					break;
				}
			}
		}

		result.add(item);

		return result;
	}

	private boolean isEmptyCell(final Cell cell)
	{
		return (cell == null) || (cell.getCellType() == Cell.CELL_TYPE_BLANK);
	}

	private boolean isCommentOutRow(final Row row)
	{
		Cell firstCell = row.getCell(0);
		if (isEmptyCell(firstCell)) {
			return false;
		}
		boolean rowIsCommentOut = (firstCell.getCellType() == Cell.CELL_TYPE_STRING)
				&& (firstCell.getStringCellValue().startsWith("#"));
		return rowIsCommentOut;
	}

	private int countNonEmptyColumns(final HSSFSheet sheet)
	{
		Row firstRow = sheet.getRow(0);
		return firstEmptyCellPosition(firstRow);
	}

	private int firstEmptyCellPosition(final Row cells)
	{
		int columnCount = 0;
		for (Cell cell : cells) {
			if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {
				break;
			}
			columnCount++;
		}
		return columnCount;
	}

	private String cellToString(final Cell cell)
	{
		String result = "";
		int type = cell.getCellType();
		if (type == Cell.CELL_TYPE_STRING) {
			result = cell.getStringCellValue().trim();
		} else if (type == Cell.CELL_TYPE_NUMERIC) {
			result = Double.toString(cell.getNumericCellValue());
			String[] s = result.split("\\.");
			if ((s.length > 1) && (Integer.valueOf(s[1]).intValue() == 0)) {
				result = s[0];
			}
		} else if (type == Cell.CELL_TYPE_BOOLEAN) {
			result = Boolean.toString(cell.getBooleanCellValue());
		}

		return result;
	}

}
