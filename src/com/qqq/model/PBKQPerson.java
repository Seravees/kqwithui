package com.qqq.model;

import java.util.List;

public class PBKQPerson {
	String name;
	String department;
	List<PBKQ> pbkqs;
	int zhongban;
	int yeban;

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getDepartment() {
		return department;
	}

	public void setDepartment(String department) {
		this.department = department;
	}

	public List<PBKQ> getPbkqs() {
		return pbkqs;
	}

	public void setPbkqs(List<PBKQ> pbkqs) {
		this.pbkqs = pbkqs;
	}

	public int getZhongban() {
		return zhongban;
	}

	public void setZhongban(int zhongban) {
		this.zhongban = zhongban;
	}

	public int getYeban() {
		return yeban;
	}

	public void setYeban(int yeban) {
		this.yeban = yeban;
	}

	@Override
	public String toString() {
		return "PBKQPerson [name=" + name + ", department=" + department
				+ ", pbkqs=" + pbkqs + ", zhongban=" + zhongban + ", yeban="
				+ yeban + "]";
	}

}
