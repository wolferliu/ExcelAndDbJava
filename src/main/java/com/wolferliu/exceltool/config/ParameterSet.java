package com.wolferliu.exceltool.config;

import com.thoughtworks.xstream.annotations.XStreamAlias;
import com.thoughtworks.xstream.annotations.XStreamOmitField;

import java.util.*;

/**
 * Created by IntelliJ IDEA.
 * User: liubing
 * Date: 3/14/11
 */
@XStreamAlias("ParameterSet")
public class ParameterSet {

	@XStreamAlias("ConnectionParameters")
	private Map<String, ConnectionParameter> connectionParameters =new HashMap<String, ConnectionParameter>();

	@XStreamAlias("SystemParameter")
	private SystemParameter systemParameter=getDefaultSystemParameter();


	public String[] getParameterNameArray(){
		List<String> retList=new ArrayList<String>();
		if(connectionParameters !=null) {
			for(String key: connectionParameters.keySet()){
				retList.add(key);
			}
		}
		String[] ret=new String[0];
		return retList.toArray(ret);
	}

	public String[] getDriverArray(){
		Map<String, String> dMap=getDriverMap();
		List<String> driverList=new ArrayList<String>();
		for(String key:dMap.keySet()){
			driverList.add(key);
		}
		return driverList.toArray(new String[]{});
	}

	@XStreamOmitField
	private Map<String, String> driverMap=null;
	private Map<String, String> getDriverMap(){
		if(driverMap==null){
			driverMap=new HashMap<String, String>();
			driverMap.put("com.mysql.jdbc.Driver",
					"jdbc:mysql://MyDbComputerIP:3306/myDatabaseName");
			driverMap.put("oracle.jdbc.driver.OracleDriver",
					"jdbc:oracle:thin:@MyDbComputerIP:1521:ORCL");
			/*
			driverMap.put("org.postgresql.Driver",
					"jdbc:postgresql://MyDbComputerIP/myDatabaseName");
			driverMap.put("com.sybase.jdbc2.jdbc.SybDriver",
					"jdbc:sybase:Tds:MyDbComputerIP:2638");
			driverMap.put("net.sourceforge.jtds.jdbc.Driver",
					"jdbc:jtds:sqlserver://MyDbComputerIP:1433/master");
			driverMap.put("com.microsoft.jdbc.sqlserver.SQLServerDriver",
					"jdbc:microsoft:sqlserver://MyDbComputerIP:1433;databaseName=master");
			driverMap.put("com.ibm.db2.jdbc.net.DB2Driver",
					"jdbc:db2://MyDbComputerIP:6789/SAMPLE");
			*/
		}
		return driverMap;
	}

	public String getDefaultUrl(String driverName){
		return getDriverMap().get(driverName);
	}

	public boolean isNameExist(String name){
		return connectionParameters.keySet().contains(name);
	}

	public void addParameter(String name, ConnectionParameter cp){
		this.connectionParameters.put(name, cp);
	}

	public ConnectionParameter getParameter(String name){
		return connectionParameters.get(name);
	}

	public ConnectionParameter getDefaultConnectionParameter(){
		ConnectionParameter cp=new ConnectionParameter();
		cp.setDataLine(6);
		cp.setTitleLine(2);
		return cp;
	}

	public SystemParameter getDefaultSystemParameter(){
		SystemParameter sp=new SystemParameter();
		sp.setCommitLines(100);
		sp.setAutoDateType(false);
		sp.setDateFormat("yyyy-MM-dd");
		sp.setTimeFormat("hh:MM:ss");
		return sp;
	}

	public void removeParameter(String name){
		connectionParameters.remove(name);
	}

	public SystemParameter getSystemParameter() {
		if(systemParameter==null){
			systemParameter=getDefaultSystemParameter();
		}
		return systemParameter;
	}

	public void setSystemParameter(SystemParameter systemParameter) {
		this.systemParameter = systemParameter;
	}
}
