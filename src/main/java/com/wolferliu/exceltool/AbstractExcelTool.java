package com.wolferliu.exceltool;

import com.wolferliu.exceltool.config.ConnectionParameter;
import com.wolferliu.exceltool.config.ParameterSet;
import com.wolferliu.exceltool.config.ParameterSetLoader;

import javax.swing.*;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.File;
import java.io.FilenameFilter;

/**
 * Created by IntelliJ IDEA.
 * Date: 3/15/11
 *
 * @author liubing
 */
public abstract class AbstractExcelTool implements AsyncUIChangable {
	protected static final String lcOSName = System.getProperty("os.name").toLowerCase();
	protected static final boolean IS_MAC = lcOSName.startsWith("mac os x");

	protected JPanel mainPanel;
	protected JFrame mainFrame;

	protected JTextArea textLog;

	//Connection Parameter
	JComboBox comboProfile;
	JComboBox comboDriver;
	JTextField textUrl;
	JTextField textUsername;
	JTextField textPassword;
	JTextField textTitleLine;
	JTextField textDataLine;
	JTextField textFileName;
	JButton buttonSelectFile;
	JButton buttonSystemConfig;

	protected ParameterSet parameterSet=new ParameterSet();
	protected ParameterSetLoader configLoader=new ParameterSetLoader();
	protected DefaultComboBoxModel profileModel;
	protected String currentProfileName;

	protected void loadConfig(){
		try {
			parameterSet=configLoader.loadConfig();
		}
		catch (Exception e) {
			e.printStackTrace();
			JOptionPane.showMessageDialog(null, "Error while loading config!", "Error", JOptionPane.ERROR_MESSAGE);
			//System.exit(-1);
		}
	}

	protected final static double labelWeight = 0.1;
	protected final static double textWeight = 1.0;

	protected void fillWithParameter(String pName) {
		ConnectionParameter cp=parameterSet.getParameter(pName);
		if(cp==null){
			cp=parameterSet.getDefaultConnectionParameter();
		}
		comboDriver.setSelectedItem(cp.getDriver());
		textUrl.setText(cp.getUrl());
		textUsername.setText(cp.getUsername());
		textDataLine.setText(String.valueOf(cp.getDataLine()));
		textTitleLine.setText(String.valueOf(cp.getTitleLine()));
	}

	protected abstract void toggleProcessing(boolean flag);

	public void startExec() {
		toggleProcessing(true);
	}

	public void endExec() {
		toggleProcessing(false);
	}

	public void log(String s) {
		textLog.append(s);
		textLog.append("\n");
	}

}
