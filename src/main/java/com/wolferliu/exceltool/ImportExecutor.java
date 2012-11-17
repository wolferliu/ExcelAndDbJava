/*
 * Created on 2004-3-5
 *
 * To change the template for this generated file go to
 * Window - Preferences - Java - Code Generation - Code and Comments
 */
package com.wolferliu.exceltool;

import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.Timestamp;
import java.util.Calendar;
import java.util.TimeZone;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

//import org.eclipse.swt.widgets.Display;
//import org.eclipse.swt.widgets.Text;


/**
 * @author SonyMusic
 *
 * To change the template for this generated type comment go to
 * Window - Preferences - Java - Code Generation - Code and Comments
 */
public final class ImportExecutor extends AbstractExcelExecutor implements Runnable {
	private String tableName;
	private int tableLine=0;
	
	//计算用字段
	String[] columnList=null;
	String[] columnTypeList=null;
	int keyCol=-1;	//第一个不为空的字段
	ExcelSheetWrapper sw=null;

	private void exec(){
		new UIChangeRunInSwing(uiObject).startExec();
		Calendar calendar=Calendar.getInstance(TimeZone.getDefault());
		Connection con = null;
		PreparedStatement ps=null;
		Workbook workbook = null;
		int readedLines=0;
		int lastRow=0;
		try {
			log("=====================================");
			log("Start import data from excel");
			log("=====================================");
			//获得连接
			//log(String.valueOf(textLog.getSize()));
			log("Connecting...");
			con = getConnection();
			con.setAutoCommit(false);
			log("Connected...");
			
			log("Opening excel file...");
			workbook = openExcelFile();
			log("Open excel file successfully...");
			
			//读取列名
			log("Loading columns...");
			loadColumnDef();
			
			//生成SQL语句
			String SQL="";
			SQL = buildInsertSQL();
			
			log("SQL is: ");
			log(SQL);
			
			
			log("Loading Data...");
			
			//按行读取数据
			ps=con.prepareStatement(SQL);

			int commitLines=systemParameter.getCommitLines();
			if(commitLines<0){
				commitLines=0;
			}
			int showLines=commitLines==0?100:commitLines;
			
			
			//Object[] dataList=new Object[sw.getTotalCols()];
			for(int row=connectionParameter.getDataLine()-1, rowCount=sw.getTotalRows(); row<rowCount; row++){
				lastRow=row;
				ps.clearParameters();
				String keyValue=sw.getString(keyCol, row, null);
				if(keyValue==null || keyValue.length()==0){
					break;
				}
				int id=1;
				//按列读取数据
				for(int col=0, colCount=sw.getTotalCols(); col<colCount; col++){
					if(columnList[col]==null){
						continue;
					}
					if("C".equalsIgnoreCase(columnTypeList[col])){
						//dataList[col]=;
						ps.setString(id++, sw.getString(col, row, null));
					}
					else if("D".equalsIgnoreCase(columnTypeList[col])){
						//dataList[col]
						java.util.Date tmpDate=sw.getDate(col, row, null);
						ps.setDate(id++, tmpDate == null ? null : new java.sql.Date(tmpDate.getTime()));
					}
					else if("T".equalsIgnoreCase(columnTypeList[col])){
						java.util.Date tmpDate=sw.getDate(col, row, null);
						ps.setTime(id++, tmpDate == null ? null : new java.sql.Time(tmpDate.getTime()));
					}
					else if("DT".equalsIgnoreCase(columnTypeList[col])){
						java.util.Date tmpDate=sw.getDate(col, row, null);
						ps.setTimestamp(id++, tmpDate == null ? null : new Timestamp(tmpDate.getTime()));
					}
					else if("I".equalsIgnoreCase(columnTypeList[col])){
						int tmpInt=sw.getInt(col, row, -1);
						ps.setInt(id++, tmpInt);
					}
					else if("N".equalsIgnoreCase(columnTypeList[col])){
						ps.setDouble(id++, sw.getDouble(col, row, -1));
					}
					else{
						throw new Exception("Error Data Type: Column: "+(col+1));
					}
				}
				
				int affectRows=ps.executeUpdate();
				if(affectRows!=1){
					throw new Exception("Data Import error! Line: "+(row+1));
				}
				readedLines++;
				if(readedLines % showLines ==0){
					if(commitLines!=0){
						con.commit();
						log("Committed "+readedLines+" lines...");
					}
					else{
						log("Loaded "+readedLines+" lines...");
					}
				}
			}
			con.commit();
			//con.rollback();
			log("Last Committed...");
			log("Import completed...");
			log("Total import "+readedLines+" lines...");
		}
		catch (Throwable e) {
			//textLog.append("\n"+e.getMessage());
			log("Error:");
			String message=e.getMessage();
			if(message==null) message="null";
			message+=" at line: "+(lastRow+1);
			log(message);
			e.printStackTrace();
			try {
				con.rollback();
			} catch (Exception ige) {
				ige.printStackTrace();
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
				if(workbook!=null) workbook.close();
			} catch (Exception ige) {
			}
		}
		new UIChangeRunInSwing(uiObject).endExec();
	}


	/**
	 * @return
	 * @throws IOException
	 * @throws BiffException
	 * @throws Exception
	 */
	private Workbook openExcelFile() throws IOException, BiffException, Exception {
		Workbook workbook;
		workbook = Workbook.getWorkbook(new java.io.File(fileName));
		Sheet sheet = workbook.getSheet(0);
		sw=new ExcelSheetWrapper(sheet);


		//若未设置tableName，则从tableLine中去读
		if(tableName==null || tableName.length()==0){

			if(tableLine==0){
				throw new Exception("You must special import table name!");
			}
			tableName=sw.getString(0, tableLine-1, null);
			if(tableName==null || tableName.length()==0){
				throw new Exception("Table name not found in excel file(CELL:A"+(tableLine)+")!");
			}
		}
		return workbook;
	}

	/**
	 * 
	 */
	private void loadColumnDef() {
		columnList=new String[sw.getTotalCols()];
		columnTypeList=new String[sw.getTotalCols()];
		int realColumnLine=connectionParameter.getTitleLine() - 1;
		for(int col=0, colCount=sw.getTotalCols(); col<colCount; col++){
			String columnName=sw.getString(col, realColumnLine, null);
			if(columnName!=null && columnName.length()>0){
				if(keyCol==-1){
					keyCol=col;
				}
				columnList[col]=columnName;
				String columnType=sw.getString(col, realColumnLine+1, null);
				if(columnType==null || columnType.length()==0){
					columnType="C";
				}
				columnTypeList[col]=columnType;
			}
		}
	}

	/**
	 * @return
	 */
	private String buildInsertSQL() {
		String SQL;
		StringBuffer buf=new StringBuffer();
		buf.append("insert into ");
		buf.append(tableName);
		buf.append(" ( ");
		int columnCount=0;
		for(int i=0; i<columnList.length; i++){
			if(columnList[i]==null || columnList[i].length()==0){
				continue;
			}
			if(columnCount>0){
				buf.append(", ");
			}
			buf.append(columnList[i]);
			columnCount++;
		}
		buf.append(") values (");
		for(int i=0; i<columnCount; i++){
			if(i>0){
				buf.append(", ");
			}
			buf.append("?");
		}
		buf.append(")");
		SQL=buf.toString();
		return SQL;
	}


	/**
	 * @return Returns the tableName.
	 */
	public String getTableName() {
		return tableName;
	}

	/**
	 * @param tableName The tableName to set.
	 */
	public void setTableName(String tableName) {
		this.tableName = tableName;
	}

	/* (non-Javadoc)
	 * @see java.lang.Runnable#run()
	 */
	public void run() {
		exec();
	}

}
