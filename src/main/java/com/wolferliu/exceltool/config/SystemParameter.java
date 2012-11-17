package com.wolferliu.exceltool.config;

import com.thoughtworks.xstream.annotations.XStreamAlias;

/**
 * Created by IntelliJ IDEA.
 * Date: 3/14/11
 *
 * @author liubing
 */
@XStreamAlias("SystemParameter")
public class SystemParameter {
	private int commitLines;
	private boolean autoDateType;
	private String dateFormat;
	private String timeFormat;

	public int getCommitLines() {
		return commitLines;
	}

	public void setCommitLines(int commitLines) {
		this.commitLines = commitLines;
	}

	public boolean isAutoDateType() {
		return autoDateType;
	}

	public void setAutoDateType(boolean autoDateType) {
		this.autoDateType = autoDateType;
	}

	public String getDateFormat() {
		return dateFormat;
	}

	public void setDateFormat(String dateFormat) {
		this.dateFormat = dateFormat;
	}

	public String getTimeFormat() {
		return timeFormat;
	}

	public void setTimeFormat(String timeFormat) {
		this.timeFormat = timeFormat;
	}
}
