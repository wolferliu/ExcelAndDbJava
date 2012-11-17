package com.wolferliu.exceltool.config;

import com.thoughtworks.xstream.annotations.XStreamAlias;
import com.thoughtworks.xstream.annotations.XStreamOmitField;

/**
 * Created by IntelliJ IDEA.
 * User: liubing
 * Date: 3/14/11
 */
@XStreamAlias("Connection")
public class ConnectionParameter {
	private String driver;
	private String url;
	private String username;

	@XStreamOmitField
	private String password;
	
	private int titleLine;
	private int dataLine;

	public String getDriver() {
		return driver;
	}

	public void setDriver(String driver) {
		this.driver = driver;
	}

	public String getUrl() {
		return url;
	}

	public void setUrl(String url) {
		this.url = url;
	}

	public String getUsername() {
		return username;
	}

	public void setUsername(String username) {
		this.username = username;
	}

	public String getPassword() {
		return password;
	}

	public void setPassword(String password) {
		this.password = password;
	}

	public int getTitleLine() {
		return titleLine;
	}

	public void setTitleLine(int titleLine) {
		this.titleLine = titleLine;
	}

	public int getDataLine() {
		return dataLine;
	}

	public void setDataLine(int dataLine) {
		this.dataLine = dataLine;
	}

}
