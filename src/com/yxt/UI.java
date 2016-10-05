package com.yxt;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;

public class UI implements ActionListener{
	private JFrame frame;
	private JPanel panel;
	private JFileChooser fileChooser;
	private JButton button;
	public UI() {
		frame = new JFrame("第一列加0处理程序（版权归杨笑天所有）");
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		panel = new JPanel();
		fileChooser = new JFileChooser();
		button = new JButton("开始处理");
		button.addActionListener(this);
		panel.add(fileChooser);
		panel.add(button);
		frame.add(panel);
		frame.setSize(520, 500);
		frame.setLocationRelativeTo(null);
		frame.setVisible(true);
	}
	@Override
	public void actionPerformed(ActionEvent e) {
		ExcelUtil.process(fileChooser.getSelectedFile().getPath());
		JOptionPane.showMessageDialog(null, "在当前路径下已生成处理好的文件yxt.xls", "提示", JOptionPane.INFORMATION_MESSAGE);
	}
}
