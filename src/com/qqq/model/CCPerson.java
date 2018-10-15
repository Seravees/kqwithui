package com.qqq.model;

import java.util.List;

public class CCPerson {
	String name;
	String department;
	List<CC> ccs;

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

	public List<CC> getCcs() {
		return ccs;
	}

	public void setCcs(List<CC> ccs) {
		this.ccs = ccs;
	}

	@Override
	public String toString() {
		return "CCPerson [name=" + name + ", department=" + department
				+ ", ccs=" + ccs + "]";
	}

}
