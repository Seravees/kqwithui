package com.qqq.model;

import java.util.List;

public class AddPerson {
	String name;
	String department;
	List<Add> adds;
	double add_weekday;
	double add_weekend;
	double add_holiday;

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

	public List<Add> getAdds() {
		return adds;
	}

	public void setAdds(List<Add> adds) {
		this.adds = adds;
	}

	public double getAdd_weekday() {
		return add_weekday;
	}

	public void setAdd_weekday(double add_weekday) {
		this.add_weekday = add_weekday;
	}

	public double getAdd_weekend() {
		return add_weekend;
	}

	public void setAdd_weekend(double add_weekend) {
		this.add_weekend = add_weekend;
	}

	public double getAdd_holiday() {
		return add_holiday;
	}

	public void setAdd_holiday(double add_holiday) {
		this.add_holiday = add_holiday;
	}

	@Override
	public String toString() {
		return "AddPerson [name=" + name + ", department=" + department
				+ ", adds=" + adds + ", add_weekday=" + add_weekday
				+ ", add_weekend=" + add_weekend + ", add_holiday="
				+ add_holiday + "]";
	}

}
