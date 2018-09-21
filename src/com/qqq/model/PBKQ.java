package com.qqq.model;

public class PBKQ {
	String name;
	String department;
	String date;
	String weekday;
	String pb;
	String strat;
	String end;
	String state;

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

	public String getDate() {
		return date;
	}

	public void setDate(String date) {
		this.date = date;
	}

	public String getWeekday() {
		return weekday;
	}

	public void setWeekday(String weekday) {
		this.weekday = weekday;
	}

	public String getPb() {
		return pb;
	}

	public void setPb(String pb) {
		this.pb = pb;
	}

	public String getStrat() {
		return strat;
	}

	public void setStrat(String strat) {
		this.strat = strat;
	}

	public String getEnd() {
		return end;
	}

	public void setEnd(String end) {
		this.end = end;
	}

	public String getState() {
		return state;
	}

	public void setState(String state) {
		this.state = state;
	}

	@Override
	public String toString() {
		return "PBKQ [name=" + name + ", department=" + department + ", date="
				+ date + ", weekday=" + weekday + ", pb=" + pb + ", strat="
				+ strat + ", end=" + end + ", state=" + state + "]";
	}

}
