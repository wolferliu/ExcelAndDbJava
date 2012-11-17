package com.wolferliu.exceltool;

/**
 * Used for write log in swing.
 * User: Liubing
 * Date: 11-2-11
 * Time: 上午12:28
 * To change this template use File | Settings | File Templates.
 */
public class LogWriterRunInSwing extends AbstractRunInSwing implements Runnable {
	String s;
	AsyncUIChangable uiObject;
	LogWriterRunInSwing(AsyncUIChangable uiObject){
		this.uiObject=uiObject;
	}
	public void run() {
		uiObject.log(s);
	}
	public void log(String s){
		this.s=s;
		runInSwing(this);
	}
}
