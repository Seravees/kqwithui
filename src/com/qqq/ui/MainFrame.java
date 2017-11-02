package com.qqq.ui;

import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.text.ParseException;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;

import com.qqq.data.Dao;
import com.qqq.model.Var;

public class MainFrame implements ActionListener {
	JPanel contentPane;

	JButton btn_start;
	File file_src;
	File file_persons;
	File file_pb1, file_pb2;
	File file_kq1, file_kq2;
	File file_out;
	File file_holiday;
	File file_add1, file_add2, file_add2se;
	File file_outpath;

	JTextField jtf_src;
	JTextField jtf_persons;
	JTextField jtf_pb1, jtf_pb2;
	JTextField jtf_kq1, jtf_kq2;
	JTextField jtf_out;
	JTextField jtf_hol;
	JTextField jtf_add1, jtf_add2, jtf_add2se;
	JTextField jtf_outpath;

	// ConsoleText showArea;
	JTextArea showArea;

	JScrollPane scrollPane;

	public MainFrame() {
		JFrame frame = new JFrame("main");

		// showArea = new ConsoleText();
		showArea = new JTextArea();
		showArea.setBounds(40, 470, 700, 70);
		showArea.setEditable(false);

		scrollPane = new JScrollPane(showArea);
		scrollPane.setBounds(40, 470, 700, 70);
		scrollPane
				.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);

		/****** 格式表格选择 ******/
		JPanel jp_src = new JPanel();
		jp_src.setLayout(new FlowLayout());
		jp_src.setBounds(0, 0, 700, 35);
		JLabel jl_src = new JLabel("源表格:");
		jl_src.setPreferredSize(new Dimension(125, 25));
		jtf_src = new JTextField();
		jtf_src.setPreferredSize(new Dimension(400, 25));
		JButton jb_src = new JButton("选择");
		jb_src.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				JFileChooser jfc = new JFileChooser();
				jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
				jfc.showDialog(new JLabel(), "选择源表格");
				file_src = jfc.getSelectedFile();
				if (file_src.isFile()) {
					jtf_src.setText(file_src.getAbsolutePath());
					// System.out.println(jtf_src.getText());
				}
			}
		});

		jp_src.add(jl_src);
		jp_src.add(jtf_src);
		jp_src.add(jb_src);

		/****** 人员名单选择 ******/
		JPanel jp_persons = new JPanel();
		jp_persons.setLayout(new FlowLayout());
		jp_persons.setBounds(0, 35, 700, 35);
		JLabel jl_persons = new JLabel("人员名单:");
		jl_persons.setPreferredSize(new Dimension(125, 25));
		jtf_persons = new JTextField();
		jtf_persons.setPreferredSize(new Dimension(400, 25));
		JButton jb_persons = new JButton("选择");
		jb_persons.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				JFileChooser jfc = new JFileChooser();
				jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
				jfc.showDialog(new JLabel(), "选择人员名单");
				file_persons = jfc.getSelectedFile();
				if (file_persons.isFile()) {
					jtf_persons.setText(file_persons.getAbsolutePath());
				}
			}
		});
		jp_persons.add(jl_persons);
		jp_persons.add(jtf_persons);
		jp_persons.add(jb_persons);

		/****** 排班表1选择 ******/
		JPanel jp_pb1 = new JPanel();
		jp_pb1.setLayout(new FlowLayout());
		jp_pb1.setBounds(0, 70, 700, 35);
		JLabel jl_pb1 = new JLabel("排班1：");
		jl_pb1.setPreferredSize(new Dimension(125, 25));
		jtf_pb1 = new JTextField();
		jtf_pb1.setPreferredSize(new Dimension(400, 25));
		JButton jb_pb1 = new JButton("选择");
		jb_pb1.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				JFileChooser jfc = new JFileChooser();
				jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
				jfc.showDialog(new JLabel(), "选择排班1");
				file_pb1 = jfc.getSelectedFile();
				if (file_pb1.isFile()) {
					jtf_pb1.setText(file_pb1.getAbsolutePath());
				}

			}
		});
		jp_pb1.add(jl_pb1);
		jp_pb1.add(jtf_pb1);
		jp_pb1.add(jb_pb1);

		/****** 排班表2选择 ******/
		JPanel jp_pb2 = new JPanel();
		jp_pb2.setLayout(new FlowLayout());
		jp_pb2.setBounds(0, 105, 700, 35);
		JLabel jl_pb2 = new JLabel("排班2：");
		jl_pb2.setPreferredSize(new Dimension(125, 25));
		jtf_pb2 = new JTextField();
		jtf_pb2.setPreferredSize(new Dimension(400, 25));
		JButton jb_pb2 = new JButton("选择");
		jb_pb2.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				JFileChooser jfc = new JFileChooser();
				jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
				jfc.showDialog(new JLabel(), "选择排班2");
				file_pb2 = jfc.getSelectedFile();
				if (file_pb2.isFile()) {
					jtf_pb2.setText(file_pb2.getAbsolutePath());
				}

			}
		});
		jp_pb2.add(jl_pb2);
		jp_pb2.add(jtf_pb2);
		jp_pb2.add(jb_pb2);

		/****** 考勤表2选择 ******/
		JPanel jp_kq1 = new JPanel();
		jp_kq1.setLayout(new FlowLayout());
		jp_kq1.setBounds(0, 140, 700, 35);
		JLabel jl_kq1 = new JLabel("考勤1：");
		jl_kq1.setPreferredSize(new Dimension(125, 25));
		jtf_kq1 = new JTextField();
		jtf_kq1.setPreferredSize(new Dimension(400, 25));
		JButton jb_kq1 = new JButton("选择");
		jb_kq1.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				JFileChooser jfc = new JFileChooser();
				jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
				jfc.showDialog(new JLabel(), "选择考勤1");
				file_kq1 = jfc.getSelectedFile();
				if (file_kq1.isFile()) {
					jtf_kq1.setText(file_kq1.getAbsolutePath());
				}

			}
		});
		jp_kq1.add(jl_kq1);
		jp_kq1.add(jtf_kq1);
		jp_kq1.add(jb_kq1);

		/****** 考勤表2选择 ******/
		JPanel jp_kq2 = new JPanel();
		jp_kq2.setLayout(new FlowLayout());
		jp_kq2.setBounds(0, 175, 700, 35);
		JLabel jl_kq2 = new JLabel("考勤2：");
		jl_kq2.setPreferredSize(new Dimension(125, 25));
		jtf_kq2 = new JTextField();
		jtf_kq2.setPreferredSize(new Dimension(400, 25));
		JButton jb_kq2 = new JButton("选择");
		jb_kq2.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				JFileChooser jfc = new JFileChooser();
				jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
				jfc.showDialog(new JLabel(), "选择考勤2");
				file_kq2 = jfc.getSelectedFile();
				if (file_kq2.isFile()) {
					jtf_kq2.setText(file_kq2.getAbsolutePath());
				}

			}
		});
		jp_kq2.add(jl_kq2);
		jp_kq2.add(jtf_kq2);
		jp_kq2.add(jb_kq2);

		/****** 外出表选择 ******/
		JPanel jp_out = new JPanel();
		jp_out.setLayout(new FlowLayout());
		jp_out.setBounds(0, 210, 700, 35);
		JLabel jl_out = new JLabel("外出:");
		jl_out.setPreferredSize(new Dimension(125, 25));
		jtf_out = new JTextField();
		jtf_out.setPreferredSize(new Dimension(400, 25));
		JButton jb_out = new JButton("选择");
		jb_out.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				JFileChooser jfc = new JFileChooser();
				jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
				jfc.showDialog(new JLabel(), "选择外出");
				file_out = jfc.getSelectedFile();
				if (file_out.isFile()) {
					jtf_out.setText(file_out.getAbsolutePath());
				}

			}
		});
		jp_out.add(jl_out);
		jp_out.add(jtf_out);
		jp_out.add(jb_out);

		/****** 请假表选择 ******/
		JPanel jp_hol = new JPanel();
		jp_hol.setLayout(new FlowLayout());
		jp_hol.setBounds(0, 245, 700, 35);
		JLabel jl_hol = new JLabel("请假:");
		jl_hol.setPreferredSize(new Dimension(125, 25));
		jtf_hol = new JTextField();
		jtf_hol.setPreferredSize(new Dimension(400, 25));
		JButton jb_hol = new JButton("选择");
		jb_hol.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				JFileChooser jfc = new JFileChooser();
				jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
				jfc.showDialog(new JLabel(), "选择请假");
				file_holiday = jfc.getSelectedFile();
				if (file_holiday.isFile()) {
					jtf_hol.setText(file_holiday.getAbsolutePath());
				}

			}
		});
		jp_hol.add(jl_hol);
		jp_hol.add(jtf_hol);
		jp_hol.add(jb_hol);

		/****** 统筹加班表选择 ******/
		JPanel jp_add1 = new JPanel();
		jp_add1.setLayout(new FlowLayout());
		jp_add1.setBounds(0, 280, 700, 35);
		JLabel jl_add1 = new JLabel("统筹加班:");
		jl_add1.setPreferredSize(new Dimension(125, 25));
		jtf_add1 = new JTextField();
		jtf_add1.setPreferredSize(new Dimension(400, 25));
		JButton jb_add1 = new JButton("选择");
		jb_add1.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				JFileChooser jfc = new JFileChooser();
				jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
				jfc.showDialog(new JLabel(), "选择统筹加班");
				file_add1 = jfc.getSelectedFile();
				if (file_add1.isFile()) {
					jtf_add1.setText(file_add1.getAbsolutePath());
				}

			}
		});
		jp_add1.add(jl_add1);
		jp_add1.add(jtf_add1);
		jp_add1.add(jb_add1);

		/****** 临时加班（完成）表选择 ******/
		JPanel jp_add2 = new JPanel();
		jp_add2.setLayout(new FlowLayout());
		jp_add2.setBounds(0, 315, 700, 35);
		JLabel jl_add2 = new JLabel("临时加班（全部）:");
		jl_add2.setPreferredSize(new Dimension(125, 25));
		jtf_add2 = new JTextField();
		jtf_add2.setPreferredSize(new Dimension(400, 25));
		JButton jb_add2 = new JButton("选择");
		jb_add2.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				JFileChooser jfc = new JFileChooser();
				jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
				jfc.showDialog(new JLabel(), "选择临时加班（全部）");
				file_add2 = jfc.getSelectedFile();
				if (file_add2.isFile()) {
					jtf_add2.setText(file_add2.getAbsolutePath());
				}

			}
		});
		jp_add2.add(jl_add2);
		jp_add2.add(jtf_add2);
		jp_add2.add(jb_add2);

		/****** 临时加班（未完成）表选择 ******/
		JPanel jp_add2se = new JPanel();
		jp_add2se.setLayout(new FlowLayout());
		jp_add2se.setBounds(0, 350, 700, 35);
		JLabel jl_add2se = new JLabel("临时加班（未完成）:");
		jl_add2se.setPreferredSize(new Dimension(125, 25));
		jtf_add2se = new JTextField();
		jtf_add2se.setPreferredSize(new Dimension(400, 25));
		JButton jb_add2se = new JButton("选择");
		jb_add2se.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				JFileChooser jfc = new JFileChooser();
				jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
				jfc.showDialog(new JLabel(), "选择临时加班（未完成）");
				file_add2se = jfc.getSelectedFile();
				if (file_add2se.isFile()) {
					jtf_add2se.setText(file_add2se.getAbsolutePath());
				}

			}
		});
		jp_add2se.add(jl_add2se);
		jp_add2se.add(jtf_add2se);
		jp_add2se.add(jb_add2se);

		/****** 输出路径选择 ******/
		JPanel jp_outpath = new JPanel();
		jp_outpath.setLayout(new FlowLayout());
		jp_outpath.setBounds(0, 385, 700, 35);
		JLabel jl_outpath = new JLabel("输出目录:");
		jl_outpath.setPreferredSize(new Dimension(125, 25));
		jtf_outpath = new JTextField();
		jtf_outpath.setPreferredSize(new Dimension(400, 25));
		JButton jb_outpath = new JButton("选择");
		jb_outpath.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				JFileChooser jfc = new JFileChooser();
				jfc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				jfc.showDialog(new JLabel(), "选择输出目录");
				file_outpath = jfc.getSelectedFile();
				if (file_outpath.isDirectory()) {
					jtf_outpath.setText(file_outpath.getAbsolutePath());
				}

			}
		});
		jp_outpath.add(jl_outpath);
		jp_outpath.add(jtf_outpath);
		jp_outpath.add(jb_outpath);

		btn_start = new JButton("START");
		btn_start.setBounds(650, 420, 75, 35);
		btn_start.addActionListener(this);

		contentPane = new JPanel();
		contentPane.setLayout(null);
		contentPane.add(jp_src);
		contentPane.add(jp_persons);
		contentPane.add(jp_pb1);
		contentPane.add(jp_pb2);
		contentPane.add(jp_kq1);
		contentPane.add(jp_kq2);
		contentPane.add(jp_out);
		contentPane.add(jp_hol);
		contentPane.add(jp_add1);
		contentPane.add(jp_add2);
		contentPane.add(jp_add2se);
		contentPane.add(jp_outpath);
		contentPane.add(btn_start);

		contentPane.add(scrollPane);
		frame.setContentPane(contentPane);
		frame.setSize(800, 600);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.setVisible(true);
		frame.setLocationRelativeTo(null);

	}

	@Override
	public void actionPerformed(ActionEvent e) {
		// TODO Auto-generated method stub
		if (e.getSource() == btn_start) {
			new Thread(new Runnable() {

				@Override
				public void run() {
					// TODO Auto-generated method stub
					showArea.setText("");

					if (!jtf_src.getText().equals("")) {
						System.out.println(file_src.getAbsolutePath());
						Var.setSrc(file_src.getAbsolutePath());
					}

					if (!jtf_persons.getText().equals("")) {
						System.out.println(file_persons.getAbsolutePath());
						Var.setName(file_persons.getAbsolutePath());
					}

					if (!jtf_pb1.getText().equals("")) {
						System.out.println(file_pb1.getAbsolutePath());
						Var.setPb1(file_pb1.getAbsolutePath());
					}

					if (!jtf_pb2.getText().equals("")) {
						System.out.println(file_pb2.getAbsolutePath());
						Var.setPb2(file_pb2.getAbsolutePath());
					} else {
						Var.setPb2("");
					}

					if (!jtf_kq1.getText().equals("")) {
						System.out.println(file_kq1.getAbsolutePath());
						Var.setKq1(file_kq1.getAbsolutePath());
					}

					if (!jtf_kq2.getText().equals("")) {
						System.out.println(file_kq2.getAbsolutePath());
						Var.setKq2(file_kq2.getAbsolutePath());
					} else {
						Var.setKq2("");
					}

					if (!jtf_out.getText().equals("")) {
						System.out.println(file_out.getAbsolutePath());
						Var.setOut(file_out.getAbsolutePath());
					} else {
						Var.setOut("");
					}

					if (!jtf_hol.getText().equals("")) {
						System.out.println(file_holiday.getAbsolutePath());
						Var.setHoliday(file_holiday.getAbsolutePath());
					} else {
						Var.setHoliday("");
					}

					if (!jtf_add1.getText().equals("")) {
						System.out.println(file_add1.getAbsolutePath());
						Var.setAdd1(file_add1.getAbsolutePath());
					} else {
						Var.setAdd1("");
					}

					if (!jtf_add2.getText().equals("")) {
						System.out.println(file_add2.getAbsolutePath());
						Var.setAdd2(file_add2.getAbsolutePath());
					} else {
						Var.setAdd2("");
					}

					if (!jtf_add2se.getText().equals("")) {
						System.out.println(file_add2se.getAbsolutePath());
						Var.setAdd2se(file_add2se.getAbsolutePath());
					} else {
						Var.setAdd2se("");
					}

					if (!jtf_outpath.getText().equals("")) {
						System.out.println(file_outpath.getAbsolutePath());
						Var.setOutPath(file_outpath.getAbsolutePath());
					}

					try {
						Dao.createPerson();
						Dao.setPB();
						Dao.setKQ();
						Dao.setOut();
						Dao.setHoliday();
						Dao.setAdd();
						Dao.fixPB();
						Dao.setRowHeight();
						Dao.merge();

						new JOptionPane();
						JOptionPane.showMessageDialog(null, "完成", "完成",
								JOptionPane.INFORMATION_MESSAGE);
					} catch (IOException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					} catch (ParseException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}

				}

			}).start();
		}
	}
}
