/*
 * Created on 2004-3-5
 *
 * To change the template for this generated file go to
 * Window - Preferences - Java - Code Generation - Code and Comments
 */
package com.wolferliu.exceltool;

import jxl.Workbook;
import jxl.biff.DisplayFormat;
import jxl.format.UnderlineStyle;
import jxl.read.biff.BiffException;
import jxl.write.*;
import jxl.write.biff.RowsExceededException;
import org.apache.commons.lang.StringUtils;

import java.io.IOException;
import java.sql.*;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.TimeZone;

//import org.eclipse.swt.widgets.Display;
//import org.eclipse.swt.widgets.Text;

/**
 * @author SonyMusic
 *
 * To change the template for this generated type comment go to
 * Window - Preferences - Java - Code Generation - Code and Comments
 */
public final class ExportExecutor extends AbstractExcelExecutor implements Runnable {

	private String SQL;

	//计算用字段
	String[] columnList=null;
	String[] columnTypeList=null;
	int[] columnWidthList=null;

	WritableSheet sheet=null;
	WritableSheet sqlSheet=null;

	private void exec(){
		new UIChangeRunInSwing(uiObject).startExec();
		Calendar calendar=Calendar.getInstance(TimeZone.getDefault());

		SimpleDateFormat dateFormat=new SimpleDateFormat(systemParameter.getDateFormat());
		SimpleDateFormat timeFormat=new SimpleDateFormat(systemParameter.getTimeFormat());
		SimpleDateFormat dateTimeFormat=new SimpleDateFormat(systemParameter.getDateFormat()
				+" "+systemParameter.getTimeFormat());
		Connection con = null;
		PreparedStatement ps=null;
		WritableWorkbook workbook = null;
		try {
			log("=====================================");
			log("Start export data to excel");
			log("=====================================");
			//获得连接
			//log(String.valueOf(textLog.getSize()));
			log("Connecting to database...");
			con = getConnection();
			con.setAutoCommit(true);
			log("Connected...");
			
			//执行SQL语句
			log("Executing SQL...");
			ps=con.prepareStatement(SQL);
			ResultSet rs=ps.executeQuery();
			ResultSetMetaData rsmd=rs.getMetaData();
			
			//取得列的名字和类型
			log("Loading columns...");
			buildColumn(rsmd);
			//logColumn();
			
			log("Creating excel file...");
			workbook = openExcelFile();
			log("Creating excel file successfully...");
			
			//向excel文件中写入SQL语句
			writeSQLToExcel();
						
			//向excel文件中写入列名
			writeColumnToExcel();
			
			log("Loading Data...");
			int readedLines=0;

			//setup cell format
			WritableFont normalFont = new WritableFont(WritableFont.ARIAL, 10, WritableFont.NO_BOLD,
					false, UnderlineStyle.NO_UNDERLINE, jxl.format.Colour.BLACK);
			WritableCellFormat cellNormalFormat = new WritableCellFormat (normalFont);

			DateFormat dfDate = new DateFormat(systemParameter.getDateFormat());
			DateFormat dfTime = new DateFormat(systemParameter.getTimeFormat());
			DateFormat dfDateTime = new DateFormat(systemParameter.getDateFormat()+" "+systemParameter.getTimeFormat());
			WritableCellFormat cellDateFormat = new WritableCellFormat (normalFont, dfDate);
			WritableCellFormat cellTimeFormat = new WritableCellFormat (normalFont, dfTime);
			WritableCellFormat cellDateTimeFormat = new WritableCellFormat (normalFont, dfDateTime);

			//按行读取数据
			int row=connectionParameter.getDataLine()-1;
			Label l=null;
			jxl.write.Number n=null;
			jxl.write.DateTime d=null;
			while(rs.next()){
				for (int i = 0; i < columnList.length; i++) {
                    int columnIndex=i+1;
					if("C".equalsIgnoreCase(columnTypeList[i])){
						l=new Label(i, row, rs.getString(columnIndex));
						sheet.addCell(l);
					}
					else if("D".equalsIgnoreCase(columnTypeList[i])){
						java.sql.Date tmpDate=rs.getDate(columnIndex);
						if(tmpDate!=null){
							//l=new Label(i, row, dateFormat.format(tmpDate));
							//sheet.addCell(l);
							d=new DateTime(i, row, tmpDate, cellDateFormat);
							sheet.addCell(d);
						}
					}
					else if("T".equalsIgnoreCase(columnTypeList[i])){
						java.sql.Time tmpDate=rs.getTime(columnIndex);
						if(tmpDate!=null){
							d=new DateTime(i, row, tmpDate, cellTimeFormat);
							sheet.addCell(d);
						}
					}
					else if("DT".equalsIgnoreCase(columnTypeList[i])){
						Timestamp tmpDate=rs.getTimestamp(columnIndex);
						if(tmpDate!=null){
							d=new DateTime(i, row, tmpDate, cellDateTimeFormat);
							sheet.addCell(d);
						}
					}
					else if("N".equalsIgnoreCase(columnTypeList[i])){
						n=new jxl.write.Number(i, row, rs.getDouble(columnIndex));
						sheet.addCell(n);
					}
					else if("I".equalsIgnoreCase(columnTypeList[i])){
						n=new jxl.write.Number(i, row, rs.getLong(columnIndex));
						sheet.addCell(n);
					}
				}
				row++;
				readedLines++;
				if(readedLines % 100 ==0){
					log("Export "+readedLines+" lines...");
				}
			}
			
			//con.commit();
			
			log("Write to excel files...");
			workbook.write();
			
			log("Export completed...");
			log("Total export "+readedLines+" lines...");
		}
		catch (Throwable e) {
			//textLog.append("\n"+e.getMessage());
			e.printStackTrace();
			log("Error:");
			String message=e.getMessage();
			if(message==null) message="null";
			log(message);
			try {
				con.rollback();
			} catch (Exception ige) {
				//ige.printStackTrace();
			}
		}
		finally{
			try {
				if(ps!=null) ps.close();
			} catch (Exception ige) {
			}
			try {
				if(con!=null) con.close();
			} catch (Exception ige) {
			}
			try {
				if(workbook!=null) {workbook.close();}
			} catch (Exception ige) {
			}
		}
		new UIChangeRunInSwing(uiObject).endExec();
	}

	private void writeSQLToExcel() throws WriteException, RowsExceededException {
		Label sqlLabel=null;
		sqlLabel=new Label(0, 0, "SQL");
		sqlSheet.addCell(sqlLabel);
		int sqlStartLine=2;
		String[] sqlArray=SQL.split("\\n");
				//BeaconString.split(SQL, "\n");
		for (int i = 0; i < sqlArray.length; i++) {
			sqlLabel=new Label(0, sqlStartLine+i, sqlArray[i]);
			sqlSheet.addCell(sqlLabel);
		}
	}
	
	
	/**
	 * @return
	 * @throws IOException
	 * @throws BiffException
	 * @throws Exception
	 */
	private WritableWorkbook openExcelFile() throws Exception {
		WritableWorkbook workbook;
		workbook = Workbook.createWorkbook(new java.io.File(fileName));
		sheet = workbook.createSheet("Export", 0);
		sqlSheet = workbook.createSheet("SQL", 1);
		
		return workbook;
	}

	/**
	 * 
	 */
	private void buildColumn(ResultSetMetaData rsmd) throws SQLException {
		columnList=new String[rsmd.getColumnCount()];
		columnTypeList=new String[rsmd.getColumnCount()];
		columnWidthList=new int[rsmd.getColumnCount()];
		int realColumnLine=connectionParameter.getTitleLine() - 1;
		for(int col=0, colCount=rsmd.getColumnCount(); col<colCount; col++){
			String columnName=rsmd.getColumnName(col+1);
			columnList[col]=columnName;
			String columnTypeName=null;
			int columnType=rsmd.getColumnType(col+1);
			System.out.println("columnType of "+columnName+" = " + columnType);
			int columnWidth=10;
			switch (columnType) {
				case java.sql.Types.CHAR :
				case java.sql.Types.LONGVARCHAR :
				case java.sql.Types.VARCHAR :
				case java.sql.Types.NCHAR :
				case java.sql.Types.LONGNVARCHAR :
				case java.sql.Types.NVARCHAR :
					columnTypeName="C";
					columnWidth=Math.max(12, Math.min(rsmd.getPrecision(col+1), 30));
					break;
				case java.sql.Types.DATE :
					columnTypeName="D";
					columnWidth=Math.max(12, columnName.length());
					break;
				case java.sql.Types.TIME :
					columnTypeName="T";
					columnWidth=Math.max(12, columnName.length());
					break;
				case java.sql.Types.TIMESTAMP :
				case -100 :
					columnTypeName="DT";
					columnWidth=Math.max(14, columnName.length());
					break;
				case java.sql.Types.NUMERIC :
				case java.sql.Types.DECIMAL :
				case java.sql.Types.INTEGER :
				case java.sql.Types.SMALLINT :
				case java.sql.Types.FLOAT :
				case java.sql.Types.REAL :
				case java.sql.Types.DOUBLE :
				case java.sql.Types.BIGINT :
				case java.sql.Types.TINYINT :
				case java.sql.Types.BIT :
					columnTypeName="N";
					if(rsmd.getScale(col+1)==0){
						columnTypeName="I";
					}
					columnWidth=Math.max(12, columnName.length());
					break;
					
				default :
					columnTypeName="C";
					break;
			}
			columnTypeList[col]=columnTypeName;
			columnWidthList[col]=columnWidth;
		}
	}
	
	private void writeColumnToExcel() throws RowsExceededException, WriteException{
		WritableFont titleFont = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD, false, UnderlineStyle.NO_UNDERLINE, jxl.format.Colour.WHITE); 
		WritableCellFormat titleFormat = new WritableCellFormat (titleFont); 
		titleFormat.setBackground(jxl.format.Colour.GRAY_80);
		Label l=null;
		
		int col=0;
		int row=connectionParameter.getTitleLine()-1;
		for (int i = 0; i < columnList.length; i++) {
			l=new Label(col+i, row, columnList[i]);
			sheet.addCell(l);
			l=new Label(col+i, row+1, columnTypeList[i]);
			sheet.addCell(l);
			l=new Label(col+i, row+3, getReadableName(columnList[i]), titleFormat);
			sheet.addCell(l);
			sheet.setColumnView(i, columnWidthList[i]);
		}
	}
	
	private String getReadableName(String colName){
		String[] names=StringUtils.split(colName, "_");
		for (int i = 0; i < names.length; i++) {
			names[i]=StringUtils.capitalize(StringUtils.lowerCase(names[i]));
		}
		String ret=StringUtils.join(names, " ");
		return ret;
	}

	/* (non-Javadoc)
	 * @see java.lang.Runnable#run()
	 */
	public void run() {
		exec();
	}

	/**
	 * @return Returns the sQL.
	 */
	public String getSQL() {
		return SQL;
	}

	/**
	 * @param sql The sQL to set.
	 */
	public void setSQL(String sql) {
		SQL = sql;
	}

}
