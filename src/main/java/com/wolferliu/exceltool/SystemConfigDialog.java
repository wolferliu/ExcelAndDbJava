package com.wolferliu.exceltool;

import com.wolferliu.exceltool.config.ParameterSet;
import com.wolferliu.exceltool.config.SystemParameter;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

/**
 * Created by IntelliJ IDEA.
 * Date: 3/14/11
 *
 * @author liubing
 */
public class SystemConfigDialog extends JDialog {
	private JTextField textCommitLines=new JTextField();
	private JCheckBox checkAutoDateType=new JCheckBox();
	private JTextField textDateFormat=new JTextField();
	private JTextField textTimeFormat=new JTextField();

	ParameterSet ps=null;
	
	public SystemConfigDialog(JFrame mainFrame, ParameterSet ps) {
		super(mainFrame);
		this.ps=ps;
		
		setTitle("System Config");
		Container cp = getContentPane();

		cp.setLayout(new BoxLayout(cp, BoxLayout.Y_AXIS));
//		GridBagConstraints c = new GridBagConstraints();
//		c.insets = new Insets(2, 2, 2, 2);
//		c.anchor = GridBagConstraints.WEST;

		cp.add(createParameterPanel());
		cp.add(Box.createVerticalStrut(50));
		cp.add(createButtonPanel());
		
		setDefaultCloseOperation(DISPOSE_ON_CLOSE);
		setSize(320, 250);
		setResizable(false);
		setModal(true);
	}

	private JPanel createParameterPanel(){
		SystemParameter sp=ps.getSystemParameter();
		
		GridBagConstraints c = new GridBagConstraints();
		c.insets = new Insets(2, 2, 2, 2);
		c.anchor = GridBagConstraints.EAST;

		JPanel paramPanel=new JPanel(new GridBagLayout());

		int x,y;

		x=0;
		y=0;
		c.weightx = 0;
		c.weighty = 0;
		c.gridx = x++;
		c.gridy = y;
		c.fill = GridBagConstraints.NONE;
		paramPanel.add(new JLabel("Commit lines:"), c);
		c.weightx = 1.0;
		c.weighty = 0;
		c.gridx = x++;
		c.gridy = y;
		c.fill = GridBagConstraints.HORIZONTAL;
		textCommitLines.setText(String.valueOf(sp.getCommitLines()));
		paramPanel.add(textCommitLines, c);

		x=0;
		y++;
		c.weightx = 0;
		c.weighty = 0;
		c.gridx = x++;
		c.gridy = y;
		c.fill = GridBagConstraints.NONE;
		paramPanel.add(new JLabel("Auto Date Type:"), c);
		c.weightx = 1.0;
		c.weighty = 0;
		c.gridx = x++;
		c.gridy = y;
		c.fill = GridBagConstraints.HORIZONTAL;
		checkAutoDateType.setSelected(sp.isAutoDateType());
		paramPanel.add(checkAutoDateType, c);

		x=0;
		y++;
		c.weightx = 0;
		c.weighty = 0;
		c.gridx = x++;
		c.gridy = y;
		c.fill = GridBagConstraints.NONE;
		paramPanel.add(new JLabel("Date Format:"), c);
		c.weightx = 1.0;
		c.weighty = 0;
		c.gridx = x++;
		c.gridy = y;
		c.fill = GridBagConstraints.HORIZONTAL;
		textDateFormat.setText(sp.getDateFormat());
		paramPanel.add(textDateFormat, c);

		x=0;
		y++;
		c.weightx = 0;
		c.weighty = 0;
		c.gridx = x++;
		c.gridy = y;
		c.fill = GridBagConstraints.NONE;
		paramPanel.add(new JLabel("Time Format:"), c);
		c.weightx = 1.0;
		c.weighty = 0;
		c.gridx = x++;
		c.gridy = y;
		c.fill = GridBagConstraints.HORIZONTAL;
		textTimeFormat.setText(sp.getTimeFormat());
		paramPanel.add(textTimeFormat, c);
		
		return paramPanel;
	}
	private JPanel createButtonPanel(){
		JPanel buttonPanel=new JPanel();
		buttonPanel.setLayout(new BoxLayout(buttonPanel, BoxLayout.LINE_AXIS));
		buttonPanel.add(Box.createHorizontalGlue());
		JButton btnOK=new JButton("OK");
		btnOK.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				SystemParameter sp=new SystemParameter();
				sp.setCommitLines(Integer.parseInt(textCommitLines.getText()));
				sp.setAutoDateType(checkAutoDateType.isSelected());
				sp.setDateFormat(textDateFormat.getText());
				sp.setTimeFormat(textTimeFormat.getText());
				ps.setSystemParameter(sp);
				SystemConfigDialog.this.dispose();
			}
		});
		buttonPanel.add(btnOK);
		JButton btnCancel=new JButton("Cancel");
		btnCancel.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				SystemConfigDialog.this.dispose();
			}
		});
		buttonPanel.add(btnCancel);
		return buttonPanel;
	}
}
