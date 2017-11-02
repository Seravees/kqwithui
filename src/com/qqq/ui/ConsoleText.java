package com.qqq.ui;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.PrintStream;

import javax.swing.JTextArea;

import com.qqq.util.LoopedStreams;

public class ConsoleText extends JTextArea {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	public ConsoleText() {
		LoopedStreams ls = null;
		try {
			ls = new LoopedStreams();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		PrintStream ps = new PrintStream(ls.getOutputStream(), true);
		System.setOut(ps);
		System.setErr(ps);
		startConsoleReaderThread(ls.getInputStream());
	}

	private void startConsoleReaderThread(InputStream inStream) {
		final BufferedReader br = new BufferedReader(new InputStreamReader(
				inStream));
		new Thread(new Runnable() {

			@Override
			public void run() {
				// TODO Auto-generated method stub
				StringBuffer sb = new StringBuffer();

				try {
					String s;
					while ((s = br.readLine()) != null) {
						sb.setLength(0);
						append(sb.append(s).append("\n").toString());
						setCaretPosition(getText().length());
					}
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		}).start();
	}
}
