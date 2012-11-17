package com.wolferliu.exceltool;

import jxl.*;
import jxl.biff.formula.FormulaException;

import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.io.File;
import java.util.Locale;
import java.util.TimeZone;

/**
 * Created by IntelliJ IDEA.
 * User: SonyMusic
 * Date: 2004-2-9
 * Time: 23:28:40
 * To change this template use Options | File Templates.
 */
public final class ExcelSheetWrapper {
	private Sheet sheet = null;
	private Cell cell = null;
	private int totalRows=0;
	private int totalCols=0;

	public ExcelSheetWrapper(Sheet sheet) {
		this.sheet = sheet;
		totalRows=sheet.getRows();
		totalCols=sheet.getColumns();
	}

	public int getTotalRows() {
		return totalRows;
	}

	public int getTotalCols() {
		return totalCols;
	}

	public String getString(int col, int row) {
		return getString(col, row, null);
	}

	public String getString(int col, int row, String defaultValue) {
		cell = sheet.getCell(col, row);
		String ret=null;
			ret = cell.getContents();
		return ret == null ? defaultValue : ret;
	}

	public double getDouble(int col, int row, double defaultValue) {
		cell = sheet.getCell(col, row);
		double ret = defaultValue;
		if (cell.getType() == CellType.NUMBER) {
			ret = ((NumberCell) cell).getValue();
		}
		else {
			try {
				ret = Double.parseDouble(getString(col, row, String.valueOf(defaultValue)));
			}
			catch (Exception e) {
				ret = defaultValue;
			}
		}
		return ret;
	}

	public int getInt(int col, int row, int defaultValue) {
		cell = sheet.getCell(col, row);
		int ret = defaultValue;
				ret = (int) getDouble(col, row, defaultValue);
		return ret;
	}

	public Date getDate(int col, int row) {
		return getDate(col, row, null);
	}

	public Date getDate(int col, int row, Date defaultValue) {
		cell = sheet.getCell(col, row);
		Date ret = null;
		if (cell.getType() == CellType.DATE) {
			ret = ((DateCell) cell).getDate();
		}
		else {
//			String tmp = cell.getContents();
//			BeaconDate temp = BeaconDate.getInstance(tmp);
//			if (temp != null) {
//				ret = temp.getDate();
//			}
		}
		try{
			return ret == null ? defaultValue : convertDate4JXL(ret);
		}
		catch (Exception e){
			return defaultValue;
		}
	}

	/**
	 * JXL中通过DateCell.getDate()获取单元格中的时间为（实际填写日期+8小时），原因是JXL是按照GMT时区来解析XML。本方法用于获取单元格中实际填写的日期！
	 * 例如单元格中日期为“2009-9-10”，getDate得到的日期便是“Thu Sep 10 08:00:00 CST 2009”；单元格中日期为“2009-9-10 16:00:00”，getDate得到的日期便是“Fri Sep 11 00:00:00 CST 2009”
	 *
	 * @param jxlDate 通过DateCell.getDate()获取的时间
	 * @return
	 * @throws ParseException
	 * @author XHY
	 */
	public static java.util.Date convertDate4JXL(java.util.Date jxlDate) throws ParseException {
		if (jxlDate == null)
			return null;
		TimeZone gmt = TimeZone.getTimeZone("GMT");
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss", Locale.getDefault());
		dateFormat.setTimeZone(gmt);
		String str = dateFormat.format(jxlDate);

		TimeZone local = TimeZone.getDefault();
		dateFormat.setTimeZone(local);
		return dateFormat.parse(str);
	}

	public String getFormula(int col, int row){
		cell = sheet.getCell(col, row);
		try {
			if(cell.getType()==CellType.NUMBER_FORMULA || cell.getType()==CellType.STRING_FORMULA){
				return ((FormulaCell) cell).getFormula();
			}
			else if(cell.getType()==CellType.FORMULA_ERROR){
				return null;
			}
			else{
				return null;
			}
		}
		catch (FormulaException e) {
			return null;
		}
	}

	public static void main(String[] args) throws Exception {
		Workbook workbook = null;

		try {
			workbook = Workbook.getWorkbook(new File("e:/priceByCustomer.xls"));
		} catch (Exception e) {
			e.printStackTrace();
			throw e;
		}
		Sheet sheet = workbook.getSheet(0);
		ExcelSheetWrapper eu=new ExcelSheetWrapper(sheet);
		System.out.println(eu.getString(5, 1, "Null"));
		System.out.println(eu.getString(6, 1, "Null"));

		System.out.println(eu.getString(1, 122, "Null"));
		System.out.println(eu.getString(1, 13, "Null"));
		System.out.println(eu.getDouble(11, 13, -99.98));
		System.out.println(eu.getDouble(11, 14, -99.98));
		System.out.println(eu.getInt(11, 13, -99));
		System.out.println(eu.getInt(11, 14, -99));
//		System.out.println(eu.getBeaconDate(9, 13));
//		System.out.println(eu.getBeaconDate(9, 14));

		System.out.println(eu.getInt(7, 7, -99));
		System.out.println(eu.getString(7, 7, "Null"));

		System.out.println(eu.getFormula(7, 7));
		System.out.println(eu.getFormula(7, 8));

		workbook.close();
	}

}
