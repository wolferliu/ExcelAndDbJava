package com.wolferliu.exceltool;

import com.wolferliu.exceltool.config.ConnectionParameter;
import com.wolferliu.exceltool.config.SystemParameter;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

/**
 * Created by IntelliJ IDEA.
 * User: Liubing
 * Date: 11-2-11
 * Time: 9:59 PM
 * 抽象层
 */
public abstract class AbstractExcelExecutor {
	protected AsyncUIChangable uiObject;
	protected LogWriterRunInSwing logWriter;

	protected ConnectionParameter connectionParameter;
	protected SystemParameter systemParameter;
	protected String fileName;

	protected void log(String s){
		logWriter.log(s);
	}

	/**
	 * @return
	 * @throws ClassNotFoundException
	 * @throws java.sql.SQLException
	 */
	protected Connection getConnection() throws ClassNotFoundException, SQLException {
		Connection con;
		Class.forName(connectionParameter.getDriver());
		con= DriverManager.getConnection(connectionParameter.getUrl(),
				connectionParameter.getUsername(), connectionParameter.getPassword());
		return con;
	}


	/**
	 * @return Returns the fileName.
	 */
	public String getFileName() {
		return fileName;
	}

	/**
	 * @param fileName The fileName to set.
	 */
	public void setFileName(String fileName) {
		this.fileName = fileName;
	}

	/**
	 * @return Returns the uiObject.
	 */
	public AsyncUIChangable getUiObject() {
		return uiObject;
	}

	/**
	 * @param uiObject The uiObject to set.
	 */
	public void setUiObject(AsyncUIChangable uiObject) {
		this.uiObject = uiObject;
		this.logWriter=new LogWriterRunInSwing(uiObject);
	}

	public ConnectionParameter getConnectionParameter() {
		return connectionParameter;
	}

	public void setConnectionParameter(ConnectionParameter connectionParameter) {
		this.connectionParameter = connectionParameter;
	}

	public SystemParameter getSystemParameter() {
		return systemParameter;
	}

	public void setSystemParameter(SystemParameter systemParameter) {
		this.systemParameter = systemParameter;
	}
}
