package com.qqq.data;

import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import com.qqq.model.Add;
import com.qqq.model.AddPerson;
import com.qqq.model.CC;
import com.qqq.model.CCPerson;
import com.qqq.model.Holiday;
import com.qqq.model.HolidayPerson;
import com.qqq.model.KQPerson;
import com.qqq.model.Out;
import com.qqq.model.OutPerson;
import com.qqq.model.PBKQ;
import com.qqq.model.PBKQPerson;
import com.qqq.model.PBPerson;
import com.qqq.model.Person;
import com.qqq.model.Var;
import com.qqq.util.Tools;

public class Dao {
	static List<Person> persons;
	static List<PBKQPerson> pbkqPersons;
	static List<PBPerson> pbPersons;
	static List<KQPerson> kqPersons;
	static List<HolidayPerson> holidayPersons;
	static List<OutPerson> outPersons;
	static List<CCPerson> ccPersons;
	static List<AddPerson> addPersons;

	public static void createPerson() throws IOException {
		System.out.println("******createPerson******");

		List<List<Object>> names = Tools.readAll(Var.getName());

		persons = new ArrayList<Person>();
		for (int i = 1; i < names.size(); i++) {
			Person person = new Person();
			person.setName((String) names.get(i).get(1));
			person.setDepartment((String) names.get(i).get(0));
			persons.add(person);
		}

		for (Person person : persons) {
			String fileName = person.getDepartment() + "-" + person.getName();
			// System.out.println(fileName);
			Workbook wb = Tools.open(Var.getSrc());
			Tools.writeString(wb, 1, 1, person.getName(), 3);
			Tools.writeString(wb, 1, 5, person.getDepartment(), 3);
			Tools.writeString(wb, 1, 10,
					new SimpleDateFormat("MM").format(new Date()), 3);

			Tools.create(Var.getOutPath(), fileName, Var.XLS, wb);
		}
	}

	public static void setPBKQ() throws IOException, ParseException {
		System.out.println("******setPBKQ******");

		List<List<Object>> pbkqs = Tools.readAll(Var.getPbkq());

		List<PBKQ> pbkqlist = new ArrayList<PBKQ>();

		for (int i = 1; i < pbkqs.size(); i++) {
			PBKQ pbkq = new PBKQ();
			pbkq.setDate(pbkqs.get(i).get(3).toString().substring(5));
			String date = pbkqs.get(i).get(3).toString();
			pbkq.setWeekday(new SimpleDateFormat("E")
					.format(new SimpleDateFormat("yyyy-MM-dd").parse(date)));
			pbkq.setDepartment(pbkqs.get(i).get(0).toString());
			pbkq.setName(pbkqs.get(i).get(2).toString());
			pbkq.setPb(pbkqs.get(i).get(4).toString());
			pbkq.setStrat(pbkqs.get(i).get(5).toString());
			pbkq.setEnd(pbkqs.get(i).get(6).toString());
			String state = pbkqs.get(i).get(8).toString();
			// System.out.println(pbkqs.get(i).get(5).toString().isEmpty());
			if (state.contains("迟到")) {
				pbkq.setStrat(pbkq.getStrat() + "(迟到)");
			}
			if (state.contains("早退")) {
				pbkq.setEnd(pbkq.getEnd() + "(早退)");
			}
			if (state.contains("旷工")) {
				if (pbkq.getStrat().isEmpty()) {
					// System.out.println("1flag");
					pbkq.setStrat("旷工");
				}
				if (pbkq.getEnd().isEmpty()) {
					// System.out.println("2flag");
					pbkq.setEnd("旷工");
				}
			}
			pbkqlist.add(pbkq);
		}

		pbkqPersons = new ArrayList<PBKQPerson>();

		for (PBKQ pbkq : pbkqlist) {
			if (pbkqPersons.size() > 0) {
				int j = 0;
				for (int i = 0; i < pbkqPersons.size(); i++) {
					if (pbkqPersons.get(i).getName().equals(pbkq.getName())
							&& pbkqPersons.get(i).getDepartment()
									.equals(pbkq.getDepartment())) {
						List<PBKQ> tempPBKQs = pbkqPersons.get(i).getPbkqs();
						tempPBKQs.add(pbkq);
						pbkqPersons.get(i).setPbkqs(tempPBKQs);
						if (pbkq.getPb().contains("中班")) {
							pbkqPersons.get(i).setZhongban(
									pbkqPersons.get(i).getZhongban() + 1);
						} else if (pbkq.getPb().contains("晚班")) {
							pbkqPersons.get(i).setYeban(
									pbkqPersons.get(i).getYeban() + 1);
						}
						j = i;
						break;
					}
				}
				if (!pbkqPersons.get(j).getName().equals(pbkq.getName())
						|| !pbkqPersons.get(j).getDepartment()
								.equals(pbkq.getDepartment())) {
					PBKQPerson tempPBKQPerson = new PBKQPerson();
					tempPBKQPerson.setDepartment(pbkq.getDepartment());
					tempPBKQPerson.setName(pbkq.getName());
					List<PBKQ> tempPBKQs = new ArrayList<PBKQ>();
					tempPBKQs.add(pbkq);
					tempPBKQPerson.setPbkqs(tempPBKQs);
					if (pbkq.getPb().contains("中班")) {
						tempPBKQPerson.setZhongban(1);
					} else if (pbkq.getPb().contains("晚班")) {
						tempPBKQPerson.setYeban(1);
					}
					pbkqPersons.add(tempPBKQPerson);
				}
			} else {
				PBKQPerson pbkqPerson = new PBKQPerson();
				pbkqPerson.setDepartment(pbkq.getDepartment());
				pbkqPerson.setName(pbkq.getName());
				List<PBKQ> tempPBKQs = new ArrayList<PBKQ>();
				tempPBKQs.add(pbkq);
				pbkqPerson.setPbkqs(tempPBKQs);
				if (pbkq.getPb().contains("中班")) {
					pbkqPerson.setZhongban(1);
				} else if (pbkq.getPb().contains("晚班")) {
					pbkqPerson.setYeban(1);
				}
				pbkqPersons.add(pbkqPerson);
			}
		}

		// for (PBKQPerson p : pbkqPersons) {
		// //System.out.println(p.toString());
		// }

		for (PBKQPerson pbkqPerson : pbkqPersons) {
			String fileName = pbkqPerson.getDepartment() + "-"
					+ pbkqPerson.getName();
			// System.out.println(fileName);
			Workbook wb = Tools.openPerson(fileName, Var.XLS);
			if (wb != null) {
				for (int i = 0; i < pbkqPerson.getPbkqs().size(); i++) {
					Tools.writeString(wb, i + 4, 0, pbkqPerson.getPbkqs()
							.get(i).getDate(), 2);
					Tools.writeString(wb, i + 4, 1, pbkqPerson.getPbkqs()
							.get(i).getWeekday(), 2);
					Tools.writeString(wb, i + 4, 2, pbkqPerson.getPbkqs()
							.get(i).getPb(), 2);
					if (!pbkqPerson.getPbkqs().get(i).getPb().equals("休息")) {
						if (pbkqPerson.getPbkqs().get(i).getStrat()
								.equals("旷工")
								|| pbkqPerson.getPbkqs().get(i).getStrat()
										.contains("迟到")
								|| pbkqPerson.getPbkqs().get(i).getStrat()
										.contains("早退")) {
							Tools.writeString(wb, i + 4, 3, pbkqPerson
									.getPbkqs().get(i).getStrat(), 1);
						} else {
							Tools.writeString(wb, i + 4, 3, pbkqPerson
									.getPbkqs().get(i).getStrat(), 2);
						}
						if (pbkqPerson.getPbkqs().get(i).getEnd().equals("旷工")
								|| pbkqPerson.getPbkqs().get(i).getEnd()
										.contains("迟到")
								|| pbkqPerson.getPbkqs().get(i).getEnd()
										.contains("早退")) {
							Tools.writeString(wb, i + 4, 4, pbkqPerson
									.getPbkqs().get(i).getEnd(), 1);
						} else {
							Tools.writeString(wb, i + 4, 4, pbkqPerson
									.getPbkqs().get(i).getEnd(), 2);
						}
					}
				}
				Sheet sheet = wb.getSheetAt(0);
				for (int i = 0; i <= sheet.getLastRowNum(); i++) {
					if (sheet.getRow(i).getCell(0).getStringCellValue()
							.equals("中班数：")) {
						Tools.writeString(wb, i, 1,
								"" + pbkqPerson.getZhongban(), 2);
					}
					if (sheet.getRow(i).getCell(3).getStringCellValue()
							.equals("夜班数:")) {
						Tools.writeString(wb, i, 4, "" + pbkqPerson.getYeban(),
								2);
					}
				}
				Tools.savePerson(fileName, Var.XLS, wb);
			}
		}
	}

	public static void setBD() throws IOException, ParseException {
		System.out.println("******setBD******");

		List<List<Object>> bds = Tools.readAll(Var.getBd());

		List<Out> outList = new ArrayList<Out>();
		List<CC> ccList = new ArrayList<CC>();
		List<Holiday> holList = new ArrayList<Holiday>();
		List<Add> addList = new ArrayList<Add>();

		// System.out.println(bds.size());

		if (bds != null) {
			for (int i = 1; i < bds.size(); i++) {
				// System.out.println(bds.get(i).get(4).toString());
				if (bds.get(i).get(4).toString().equals("外出")) {
					Out out = new Out();
					out.setDepartment(bds.get(i).get(0).toString());
					out.setName(bds.get(i).get(2).toString());
					out.setDate(bds.get(i).get(5).toString().substring(5, 10));
					String start = bds.get(i).get(5).toString().substring(11);
					String end = bds.get(i).get(6).toString().substring(11);
					SimpleDateFormat sdf = new SimpleDateFormat("HH:mm");
					Date start1 = sdf.parse(start);
					Date end1 = sdf.parse(end);
					if (start1.before(sdf.parse("09:01"))) {
						out.setStart("外出");
					} else {
						out.setStart(" ");
					}
					if (end1.after(sdf.parse("17:29"))) {
						out.setEnd("外出");
					} else {
						out.setEnd(" ");
					}
					out.setState(bds.get(i).get(10).toString());

					outList.add(out);
				}

				if (bds.get(i).get(4).toString().equals("出差")) {
					CC cc = new CC();
					cc.setDepartment(bds.get(i).get(0).toString());
					cc.setName(bds.get(i).get(2).toString());
					cc.setStart(bds.get(i).get(5).toString().substring(5));
					cc.setEnd(bds.get(i).get(6).toString().substring(5));
					cc.setState(bds.get(i).get(10).toString());

					ccList.add(cc);
				}

				if (bds.get(i).get(3).toString().equals("请假")) {
					Holiday hol = new Holiday();
					hol.setDepartment(bds.get(i).get(0).toString());
					hol.setName(bds.get(i).get(2).toString());
					hol.setStart(bds.get(i).get(5).toString());
					hol.setEnd(bds.get(i).get(6).toString());
					hol.setType(bds.get(i).get(4).toString());
					String hour = bds.get(i).get(7).toString();
					double hours = 0;
					if (hour.contains("天")) {
						if (hour.contains("小时")) {
							hours = Double.parseDouble(hour.substring(0,
									hour.indexOf("天")))
									* 8
									+ Double.parseDouble(hour.substring(
											hour.indexOf("天") + 1,
											hour.indexOf("小时")));
						} else {
							hours = Double.parseDouble(hour.substring(0,
									hour.indexOf("天"))) * 8;
						}
					} else if (hour.contains("小时")) {
						hours = Double.parseDouble(hour.substring(0,
								hour.indexOf("小时")));
					}
					hol.setHours(hours);

					holList.add(hol);
				}

				if (bds.get(i).get(3).toString().equals("加班")) {
					Add add = new Add();
					add.setDepartment(bds.get(i).get(0).toString());
					add.setName(bds.get(i).get(2).toString());
					String start = bds.get(i).get(5).toString();
					String end = bds.get(i).get(6).toString();
					add.setType(bds.get(i).get(4).toString());
					add.setStart(start);
					add.setEnd(end);
					add.setSite(bds.get(i).get(8).toString());
					// SimpleDateFormat sdf = new SimpleDateFormat(
					// "yyyy-MM-dd HH:mm");
					// double hours = (double)(sdf.parse(end).getTime() -
					// sdf.parse(start)
					// .getTime()) / 60 / 60 / 1000;
					String hour = bds.get(i).get(7).toString();
					double hours = Double.parseDouble(hour.substring(0,
							hour.indexOf("小时")));
					add.setHours(hours);

					addList.add(add);
				}
			}
		}

		outPersons = new ArrayList<OutPerson>();

		for (Out out : outList) {
			if (outPersons.size() > 0) {
				int j = 0;
				for (int i = 0; i < outPersons.size(); i++) {
					if (outPersons.get(i).getName().equals(out.getName())
							&& outPersons.get(i).getDepartment()
									.equals(out.getDepartment())) {
						List<Out> tempOuts = outPersons.get(i).getOuts();
						tempOuts.add(out);
						outPersons.get(i).setOuts(tempOuts);
						j = i;
						break;
					}
				}
				if (!outPersons.get(j).getName().equals(out.getName())
						|| !outPersons.get(j).getDepartment()
								.equals(out.getDepartment())) {
					OutPerson tempOutPerson = new OutPerson();
					tempOutPerson.setName(out.getName());
					tempOutPerson.setDepartment(out.getDepartment());
					List<Out> tempOuts = new ArrayList<Out>();
					tempOuts.add(out);
					tempOutPerson.setOuts(tempOuts);
					outPersons.add(tempOutPerson);
				}
			} else {
				OutPerson tempOutPerson = new OutPerson();
				tempOutPerson.setName(out.getName());
				tempOutPerson.setDepartment(out.getDepartment());
				List<Out> tempOuts = new ArrayList<Out>();
				tempOuts.add(out);
				tempOutPerson.setOuts(tempOuts);
				outPersons.add(tempOutPerson);
			}
		}

		ccPersons = new ArrayList<CCPerson>();

		for (CC cc : ccList) {
			if (ccPersons.size() > 0) {
				int j = 0;
				for (int i = 0; i < ccPersons.size(); i++) {
					if (ccPersons.get(i).getName().equals(cc.getName())
							&& ccPersons.get(i).getDepartment()
									.equals(cc.getDepartment())) {
						List<CC> tempCCs = ccPersons.get(i).getCcs();
						tempCCs.add(cc);
						ccPersons.get(i).setCcs(tempCCs);
						j = i;
						break;
					}
				}
				if (!ccPersons.get(j).getName().equals(cc.getName())
						|| !ccPersons.get(j).getDepartment()
								.equals(cc.getDepartment())) {
					CCPerson tempCCPerson = new CCPerson();
					tempCCPerson.setName(cc.getName());
					tempCCPerson.setDepartment(cc.getDepartment());
					List<CC> tempCCs = new ArrayList<CC>();
					tempCCs.add(cc);
					tempCCPerson.setCcs(tempCCs);
					ccPersons.add(tempCCPerson);
				}
			} else {
				CCPerson tempCCPerson = new CCPerson();
				tempCCPerson.setName(cc.getName());
				tempCCPerson.setDepartment(cc.getDepartment());
				List<CC> tempCCs = new ArrayList<CC>();
				tempCCs.add(cc);
				tempCCPerson.setCcs(tempCCs);
				ccPersons.add(tempCCPerson);
			}
		}

		holidayPersons = new ArrayList<HolidayPerson>();

		for (Holiday hol : holList) {
			if (holidayPersons.size() > 0) {
				int j = 0;
				for (int i = 0; i < holidayPersons.size(); i++) {
					if (holidayPersons.get(i).getName().equals(hol.getName())
							&& holidayPersons.get(i).getDepartment()
									.equals(hol.getDepartment())) {
						List<Holiday> tempHolidays = holidayPersons.get(i)
								.getHolidays();
						tempHolidays.add(hol);
						holidayPersons.get(i).setHolidays(tempHolidays);
						if (hol.getType().equals("年假")) {
							holidayPersons.get(i).setNianjia(
									holidayPersons.get(i).getNianjia()
											+ hol.getHours());
						} else if (hol.getType().equals("门诊病假")) {
							holidayPersons.get(i).setBingjia(
									holidayPersons.get(i).getBingjia()
											+ hol.getHours());
						} else if (hol.getType().equals("事假")) {
							holidayPersons.get(i).setShijia(
									holidayPersons.get(i).getShijia()
											+ hol.getHours());
						} else if (hol.getType().equals("调休")) {
							holidayPersons.get(i).setTiaoxiu(
									holidayPersons.get(i).getTiaoxiu()
											+ hol.getHours());
						} else {
							holidayPersons.get(i).setQita(
									holidayPersons.get(i).getQita()
											+ hol.getHours());
						}
						j = i;
						break;
					}
				}
				if (!holidayPersons.get(j).getName().equals(hol.getName())
						|| !holidayPersons.get(j).getDepartment()
								.equals(hol.getDepartment())) {
					HolidayPerson tempHolidayPerson = new HolidayPerson();
					tempHolidayPerson.setName(hol.getName());
					tempHolidayPerson.setDepartment(hol.getDepartment());
					List<Holiday> tempHolidays = new ArrayList<Holiday>();
					tempHolidays.add(hol);
					tempHolidayPerson.setHolidays(tempHolidays);
					if (hol.getType().equals("年假")) {
						tempHolidayPerson.setNianjia(hol.getHours());
					} else if (hol.getType().equals("门诊病假")) {
						tempHolidayPerson.setBingjia(hol.getHours());
					} else if (hol.getType().equals("事假")) {
						tempHolidayPerson.setShijia(hol.getHours());
					} else if (hol.getType().equals("调休")) {
						tempHolidayPerson.setTiaoxiu(hol.getHours());
					} else {
						tempHolidayPerson.setQita(hol.getHours());
					}
					holidayPersons.add(tempHolidayPerson);
				}
			} else {
				HolidayPerson tempHolidayPerson = new HolidayPerson();
				tempHolidayPerson.setName(hol.getName());
				tempHolidayPerson.setDepartment(hol.getDepartment());
				List<Holiday> tempHolidays = new ArrayList<Holiday>();
				tempHolidays.add(hol);
				tempHolidayPerson.setHolidays(tempHolidays);
				if (hol.getType().equals("年假")) {
					tempHolidayPerson.setNianjia(hol.getHours());
				} else if (hol.getType().equals("门诊病假")) {
					tempHolidayPerson.setBingjia(hol.getHours());
				} else if (hol.getType().equals("事假")) {
					tempHolidayPerson.setShijia(hol.getHours());
				} else if (hol.getType().equals("调休")) {
					tempHolidayPerson.setTiaoxiu(hol.getHours());
				} else {
					tempHolidayPerson.setQita(hol.getHours());
				}
				holidayPersons.add(tempHolidayPerson);
			}
		}

		addPersons = new ArrayList<AddPerson>();

		for (Add add : addList) {
			if (addPersons.size() > 0) {
				int j = 0;
				for (int i = 0; i < addPersons.size(); i++) {
					if (addPersons.get(i).getName().equals(add.getName())
							&& addPersons.get(i).getDepartment()
									.equals(add.getDepartment())) {
						List<Add> tempAdds = addPersons.get(i).getAdds();
						tempAdds.add(add);
						addPersons.get(i).setAdds(tempAdds);
						if (add.getType().equals("工作日加班")) {
							addPersons.get(i).setAdd_weekday(
									addPersons.get(i).getAdd_weekday()
											+ add.getHours());
						} else if (add.getType().equals("休息日加班")) {
							addPersons.get(i).setAdd_weekend(
									addPersons.get(i).getAdd_weekend()
											+ add.getHours());
						} else if (add.getType().equals("节假日加班")) {
							addPersons.get(i).setAdd_holiday(
									addPersons.get(i).getAdd_holiday()
											+ add.getHours());
						}
						j = i;
						break;
					}
				}
				if (!addPersons.get(j).getName().equals(add.getName())
						|| !addPersons.get(j).getDepartment()
								.equals(add.getDepartment())) {
					AddPerson tempAddPerson = new AddPerson();
					tempAddPerson.setName(add.getName());
					tempAddPerson.setDepartment(add.getDepartment());
					List<Add> tempAdds = new ArrayList<Add>();
					tempAdds.add(add);
					tempAddPerson.setAdds(tempAdds);
					if (add.getType().equals("工作日加班")) {
						tempAddPerson.setAdd_weekday(add.getHours());
					} else if (add.getType().equals("休息日加班")) {
						tempAddPerson.setAdd_weekend(add.getHours());
					} else if (add.getType().equals("节假日加班")) {
						tempAddPerson.setAdd_holiday(add.getHours());
					}
					addPersons.add(tempAddPerson);
				}
			} else {
				AddPerson tempAddPerson = new AddPerson();
				tempAddPerson.setName(add.getName());
				tempAddPerson.setDepartment(add.getDepartment());
				List<Add> tempAdds = new ArrayList<Add>();
				tempAdds.add(add);
				tempAddPerson.setAdds(tempAdds);
				if (add.getType().equals("工作日加班")) {
					tempAddPerson.setAdd_weekday(add.getHours());
				} else if (add.getType().equals("休息日加班")) {
					tempAddPerson.setAdd_weekend(add.getHours());
				} else if (add.getType().equals("节假日加班")) {
					tempAddPerson.setAdd_holiday(add.getHours());
				}
				addPersons.add(tempAddPerson);
			}
		}

		for (OutPerson outPerson : outPersons) {
			String fileName = outPerson.getDepartment() + "-"
					+ outPerson.getName();
			System.out.println(fileName);
			Workbook wb = Tools.openPerson(fileName, Var.XLS);
			for (Out out : outPerson.getOuts()) {
				Sheet sheet = wb.getSheetAt(0);
				String date = out.getDate();
				for (int i = 4; i < sheet.getLastRowNum(); i++) {
					if (sheet.getRow(i).getCell(0).getStringCellValue()
							.equals(date)) {
						if ((sheet.getRow(i).getCell(3).getStringCellValue()
								.equals("旷工") || sheet.getRow(i).getCell(3)
								.getStringCellValue().isEmpty())
								&& out.getStart().equals("外出")) {
							Tools.writeString(wb, i, 3, "外出", 2);
						}
						System.out.println();
						if ((sheet.getRow(i).getCell(4).getStringCellValue()
								.equals("旷工") || sheet.getRow(i).getCell(4)
								.getStringCellValue().isEmpty())
								&& out.getEnd().equals("外出")) {
							Tools.writeString(wb, i, 4, "外出", 2);
						}
					}
				}
			}
			Tools.savePerson(fileName, Var.XLS, wb);
		}

		for (CCPerson ccPerson : ccPersons) {
			String fileName = ccPerson.getDepartment() + "-"
					+ ccPerson.getName();
			System.out.println(fileName);
			Workbook wb = Tools.openPerson(fileName, Var.XLS);
			for (CC cc : ccPerson.getCcs()) {
				Sheet sheet = wb.getSheetAt(0);
				for (int i = 4; i < sheet.getLastRowNum(); i++) {
					String date = sheet.getRow(i).getCell(0)
							.getStringCellValue();
					SimpleDateFormat sdf = new SimpleDateFormat("MM-dd");
					if (date.contains("-")) {
						if (!sdf.parse(date).before(sdf.parse(cc.getStart()))
								&& !sdf.parse(date).after(
										sdf.parse(cc.getEnd()))) {
							// System.out.println("flag");
							if (sheet.getRow(i).getCell(3).getStringCellValue()
									.equals("旷工")
									|| sheet.getRow(i).getCell(3)
											.getStringCellValue().isEmpty()) {
								Tools.writeString(wb, i, 3, "出差", 2);
							}
							if (sheet.getRow(i).getCell(4).getStringCellValue()
									.equals("旷工")
									|| sheet.getRow(i).getCell(4)
											.getStringCellValue().isEmpty()) {
								Tools.writeString(wb, i, 4, "出差", 2);
							}
						}
					}
				}
			}
			Tools.savePerson(fileName, Var.XLS, wb);
		}

		for (HolidayPerson holidayPerson : holidayPersons) {
			String fileName = holidayPerson.getDepartment() + "-"
					+ holidayPerson.getName();
			System.out.println(fileName);
			Workbook wb = Tools.openPerson(fileName, Var.XLS);
			Sheet sheet = wb.getSheetAt(0);
			// for (int i = 36; i < sheet.getLastRowNum(); i++) {
			// if (sheet.getRow(i).getCell(0).getStringCellValue()
			// .equals("年假：")) {
			// Tools.writeDouble(wb, i, 1, holidayPerson.getNianjia(),
			// false);
			// Tools.writeDouble(wb, i, 4, holidayPerson.getBingjia(),
			// false);
			// Tools.writeDouble(wb, i, 7, holidayPerson.getShijia(),
			// false);
			// Tools.writeDouble(wb, i, 10, holidayPerson.getTiaoxiu(),
			// false);
			// Tools.writeDouble(wb, i, 13, holidayPerson.getQita(), false);
			// }
			// }
			for (Holiday holiday : holidayPerson.getHolidays()) {
				// Sheet sheet = wb.getSheetAt(0);
				for (int i = 4; i < sheet.getLastRowNum(); i++) {
					String date = sheet.getRow(i).getCell(0)
							.getStringCellValue();
					SimpleDateFormat sdf = new SimpleDateFormat("MM-dd");
					if (date.contains("-")) {
						if (!sdf.parse(date).before(
								sdf.parse(holiday.getStart().substring(5, 10)))
								&& !sdf.parse(date).after(
										sdf.parse(holiday.getEnd().substring(5,
												10)))) {
							// System.out.println("flag");
							if (sheet.getRow(i).getCell(3).getStringCellValue()
									.equals("旷工")
									|| sheet.getRow(i).getCell(3)
											.getStringCellValue().isEmpty()) {
								Tools.writeString(wb, i, 3,
										"请假：" + holiday.getType(), 2);
							}
							if (sheet.getRow(i).getCell(4).getStringCellValue()
									.equals("旷工")
									|| sheet.getRow(i).getCell(4)
											.getStringCellValue().isEmpty()) {
								Tools.writeString(wb, i, 4,
										"请假：" + holiday.getType(), 2);
							}
						}
					}

				}
				for (int i = 4; i < sheet.getLastRowNum(); i++) {
					String date = holiday.getStart().substring(5, 10);
					if (sheet.getRow(i).getCell(0).getStringCellValue()
							.equals(date)) {
						if (sheet.getRow(i).getCell(10).getStringCellValue()
								.equals("")
								|| sheet.getRow(i).getCell(10) == null) {
							Tools.writeString(wb, i, 10, holiday.getStart()
									.substring(5), 2);
							Tools.writeString(wb, i, 11, holiday.getEnd()
									.substring(5), 2);
							Tools.writeDouble(wb, i, 12, holiday.getHours(),
									true);
							Tools.writeString(wb, i, 13, holiday.getType(), 2);
							break;
						} else if (!sheet.getRow(i + 1).getCell(0)
								.getStringCellValue().equals(date)) {
							// System.out.println("" + i);
							Tools.shift(wb, i + 1);
							Tools.writeString(wb, i + 1, 10, holiday.getStart()
									.substring(5), 2);
							Tools.writeString(wb, i + 1, 11, holiday.getEnd()
									.substring(5), 2);
							Tools.writeDouble(wb, i + 1, 12,
									holiday.getHours(), true);
							Tools.writeString(wb, i + 1, 13, holiday.getType(),
									2);
							break;
						}
					}
				}
			}
			Tools.savePerson(fileName, Var.XLS, wb);
		}

		for (AddPerson addPerson : addPersons) {
			String fileName = addPerson.getDepartment() + "-"
					+ addPerson.getName();
			System.out.println(fileName);
			Workbook wb = Tools.openPerson(fileName, Var.XLS);
			Sheet sheet = wb.getSheetAt(0);
			// for (int i = 36; i < sheet.getLastRowNum(); i++) {
			// if (sheet.getRow(i).getCell(0).getStringCellValue()
			// .equals("中班数：")) {
			// Tools.writeDouble(wb, i, 7, addPerson.getAdd_weekday(),
			// false);
			// Tools.writeDouble(wb, i, 10, addPerson.getAdd_weekend(),
			// false);
			// Tools.writeDouble(wb, i, 13, addPerson.getAdd_holiday(),
			// false);
			// }
			// }
			for (Add add : addPerson.getAdds()) {
				// Sheet sheet = wb.getSheetAt(0);
				for (int i = 4; i < sheet.getLastRowNum(); i++) {
					String date = add.getStart().substring(5, 10);
					if (sheet.getRow(i).getCell(0).getStringCellValue()
							.equals(date)) {
						if (sheet.getRow(i).getCell(6).getStringCellValue()
								.equals("")
								|| sheet.getRow(i).getCell(6) == null) {
							Tools.writeString(wb, i, 5, add.getSite(), 2);
							Tools.writeString(wb, i, 6, add.getStart()
									.substring(5), 2);
							Tools.writeString(wb, i, 7,
									add.getEnd().substring(5), 2);
							Tools.writeDouble(wb, i, 8, add.getHours(), true);
							break;
						} else if (!sheet.getRow(i + 1).getCell(0)
								.getStringCellValue().equals(date)) {
							System.out.println(sheet.getRow(i).getCell(6)
									.getStringCellValue());
							if (sheet.getRow(i).getCell(6).getStringCellValue()
									.equals(add.getStart().substring(5))
									&& sheet.getRow(i).getCell(7)
											.getStringCellValue()
											.equals(add.getEnd().substring(5))) {
								if (sheet.getRow(i).getCell(5)
										.getStringCellValue().isEmpty()
										|| sheet.getRow(i).getCell(5) == null) {
									Tools.writeString(wb, i, 5, add.getSite(),
											2);
									break;
								} else {
									break;
								}
							} else {
								Tools.shift(wb, i + 1);
								Tools.writeString(wb, i + 1, 5, add.getSite(),
										2);
								Tools.writeString(wb, i + 1, 6, add.getStart()
										.substring(5), 2);
								Tools.writeString(wb, i + 1, 7, add.getEnd()
										.substring(5), 2);
								Tools.writeDouble(wb, i + 1, 8, add.getHours(),
										true);
								break;
							}
						}
					}
				}
			}
			Tools.savePerson(fileName, Var.XLS, wb);
		}
	}

	// public static void setPB() throws IOException, ParseException {
	// System.out.println("******setPB******");
	//
	// List<List<Object>> pb1s = Tools.readAll(Var.getPb1());
	// List<List<Object>> pb2s = Tools.readAll(Var.getPb2());
	//
	// List<PB> pbs = new ArrayList<PB>();
	//
	// int temp1 = 03;
	// String dateStart = new SimpleDateFormat("MM-dd")
	// .format(new SimpleDateFormat("MM.dd").parse(Var.getPb1().split(
	// "-")[0].substring(Var.getPb1().split("-")[0].length() - 4)));
	// for (int i = 3; i < pb1s.get(0).size(); i++) {
	// if (pb1s.get(0).get(i).toString().substring(0, 5).equals(dateStart)) {
	// temp1 = i;
	// }
	// }
	//
	// for (int i = 1; i < pb1s.size(); i++) {
	// for (int j = temp1; j < pb1s.get(i).size(); j++) {
	// PB pb = new PB();
	// pb.setDate(pb1s.get(0).get(j).toString().substring(0, 5));
	// pb.setWeekday(pb1s.get(0).get(j).toString().substring(6, 9));
	// pb.setDepartment(pb1s.get(i).get(0).toString());
	// pb.setName(pb1s.get(i).get(2).toString());
	// String tempPb = pb1s.get(i).get(j).toString();
	// if (tempPb.length() > 0) {
	// pb.setPb(tempPb.substring(0, tempPb.length() - 1));
	// } else {
	// pb.setPb(tempPb);
	// }
	// pbs.add(pb);
	// }
	// }
	// if (pb2s != null) {
	// int temp2 = pb2s.size();
	// String dateEnd = new SimpleDateFormat("MM-dd")
	// .format(new SimpleDateFormat("MM.dd").parse(Var.getPb2()
	// .split("-")[1].substring(0,
	// Var.getPb2().split("-")[1].lastIndexOf("."))));
	// for (int i = 3; i < pb2s.get(0).size(); i++) {
	// if (pb2s.get(0).get(i).toString().substring(0, 5)
	// .equals(dateEnd)) {
	// temp2 = i + 1;
	// }
	// }
	//
	// for (int i = 1; i < pb2s.size(); i++) {
	// int tempend = temp2 < pb2s.get(i).size() ? temp2 : pb2s.get(i)
	// .size();
	// for (int j = 3; j < tempend; j++) {
	// PB pb = new PB();
	// pb.setDate(pb2s.get(0).get(j).toString().substring(0, 5));
	// pb.setWeekday(pb2s.get(0).get(j).toString().substring(6, 9));
	// pb.setDepartment(pb2s.get(i).get(0).toString());
	// pb.setName(pb2s.get(i).get(2).toString());
	// String tempPb = pb2s.get(i).get(j).toString();
	// if (tempPb.length() > 0) {
	// pb.setPb(tempPb.substring(0, tempPb.length() - 1));
	// } else {
	// pb.setPb(tempPb);
	// }
	// pbs.add(pb);
	// }
	// }
	// }
	// pbPersons = new ArrayList<PBPerson>();
	//
	// for (PB pb : pbs) {
	// if (pbPersons.size() > 0) {
	// int j = 0;
	// for (int i = 0; i < pbPersons.size(); i++) {
	// if (pbPersons.get(i).getName().equals(pb.getName())
	// && pbPersons.get(i).getDepartment()
	// .equals(pb.getDepartment())) {
	// List<PB> tempPBs = pbPersons.get(i).getPbs();
	// tempPBs.add(pb);
	// pbPersons.get(i).setPbs(tempPBs);
	// if (pb.getPb().contains("中班")) {
	// pbPersons.get(i).setZhongban(
	// pbPersons.get(i).getZhongban() + 1);
	// } else if (pb.getPb().contains("晚班")) {
	// pbPersons.get(i).setYeban(
	// pbPersons.get(i).getYeban() + 1);
	// }
	// j = i;
	// break;
	// }
	// }
	// if (!pbPersons.get(j).getName().equals(pb.getName())
	// || !pbPersons.get(j).getDepartment()
	// .equals(pb.getDepartment())) {
	// PBPerson tempPBPerson = new PBPerson();
	// tempPBPerson.setDepartment(pb.getDepartment());
	// tempPBPerson.setName(pb.getName());
	// List<PB> tempPBs = new ArrayList<PB>();
	// tempPBs.add(pb);
	// tempPBPerson.setPbs(tempPBs);
	// if (pb.getPb().contains("中班")) {
	// tempPBPerson.setZhongban(1);
	// } else if (pb.getPb().contains("晚班")) {
	// tempPBPerson.setYeban(1);
	// }
	// pbPersons.add(tempPBPerson);
	// }
	// } else {
	// PBPerson pbPerson = new PBPerson();
	// pbPerson.setDepartment(pb.getDepartment());
	// pbPerson.setName(pb.getName());
	// List<PB> tempPBs = new ArrayList<PB>();
	// tempPBs.add(pb);
	// pbPerson.setPbs(tempPBs);
	// if (pb.getPb().contains("中班")) {
	// pbPerson.setZhongban(1);
	// } else if (pb.getPb().contains("晚班")) {
	// pbPerson.setYeban(1);
	// }
	// pbPersons.add(pbPerson);
	// }
	// }
	//
	// for (PBPerson pbPerson : pbPersons) {
	// String fileName = pbPerson.getDepartment() + "-"
	// + pbPerson.getName();
	// // System.out.println(fileName);
	// Workbook wb = Tools.openPerson(fileName, Var.XLS);
	// for (int i = 0; i < pbPerson.getPbs().size(); i++) {
	// Tools.writeString(wb, i + 4, 0, pbPerson.getPbs().get(i)
	// .getDate(), true);
	// Tools.writeString(wb, i + 4, 1, pbPerson.getPbs().get(i)
	// .getWeekday(), true);
	// Tools.writeString(wb, i + 4, 2, pbPerson.getPbs().get(i)
	// .getPb(), true);
	// }
	// Sheet sheet = wb.getSheetAt(0);
	// for (int i = 0; i <= sheet.getLastRowNum(); i++) {
	// if (sheet.getRow(i).getCell(0).getStringCellValue()
	// .equals("中班数：")) {
	// Tools.writeString(wb, i, 1, "" + pbPerson.getZhongban(),
	// false);
	// }
	// if (sheet.getRow(i).getCell(3).getStringCellValue()
	// .equals("夜班数:")) {
	// Tools.writeString(wb, i, 4, "" + pbPerson.getYeban(), false);
	// }
	// }
	// Tools.savePerson(fileName, Var.XLS, wb);
	// }
	// }

	// public static void setKQ() throws IOException {
	// System.out.println("******setKQ******");
	//
	// List<List<Object>> kq1s = Tools.readAll(Var.getKq1());
	// List<List<Object>> kq2s = Tools.readAll(Var.getKq2());
	//
	// List<KQ> kqs = new ArrayList<KQ>();
	//
	// for (int i = 2; i < kq1s.size(); i++) {
	// for (int j = 1; j < kq1s.get(i).size() / 2; j++) {
	// KQ kq = new KQ();
	// kq.setDate(kq1s.get(0).get(j * 2).toString().substring(0, 5));
	// kq.setWeekday(kq1s.get(0).get(j * 2).toString().substring(6, 9));
	// kq.setDepartment(kq1s.get(i).get(0).toString());
	// kq.setName(kq1s.get(i).get(1).toString());
	// String tempStart = kq1s.get(i).get(j * 2).toString();
	// if (tempStart.length() > 0) {
	// kq.setStart(tempStart.substring(0, tempStart.length() - 1));
	// } else {
	// kq.setStart(tempStart);
	// }
	// String tempEnd = kq1s.get(i).get(j * 2 + 1).toString();
	// if (tempEnd.length() > 0) {
	// kq.setEnd(tempEnd.substring(0, tempEnd.length() - 1));
	// } else {
	// kq.setEnd(tempEnd);
	//
	// }
	// kqs.add(kq);
	// }
	// }
	//
	// if (kq2s != null) {
	// for (int i = 2; i < kq2s.size(); i++) {
	// for (int j = 1; j < kq2s.get(i).size() / 2; j++) {
	// KQ kq = new KQ();
	// kq.setDate(kq2s.get(0).get(j * 2).toString()
	// .substring(0, 5));
	// kq.setWeekday(kq2s.get(0).get(j * 2).toString()
	// .substring(6, 9));
	// kq.setDepartment(kq2s.get(i).get(0).toString());
	// kq.setName(kq2s.get(i).get(1).toString());
	// String tempStart = kq2s.get(i).get(j * 2).toString();
	// if (tempStart.length() > 0) {
	// kq.setStart(tempStart.substring(0,
	// tempStart.length() - 1));
	// } else {
	// kq.setStart(tempStart);
	// }
	// String tempEnd = kq2s.get(i).get(j * 2 + 1).toString();
	// if (tempEnd.length() > 0) {
	// kq.setEnd(tempEnd.substring(0, tempEnd.length() - 1));
	// } else {
	// kq.setEnd(tempEnd);
	//
	// }
	// kqs.add(kq);
	// }
	// }
	// }
	// kqPersons = new ArrayList<KQPerson>();
	//
	// for (KQ kq : kqs) {
	// if (kqPersons.size() > 0) {
	// int j = 0;
	// for (int i = 0; i < kqPersons.size(); i++) {
	// if (kqPersons.get(i).getDepartment()
	// .equals(kq.getDepartment())
	// && kqPersons.get(i).getName().equals(kq.getName())) {
	// List<KQ> tempKQs = kqPersons.get(i).getKqs();
	// tempKQs.add(kq);
	// kqPersons.get(i).setKqs(tempKQs);
	// j = i;
	// break;
	// }
	// }
	// if (!kqPersons.get(j).getDepartment()
	// .equals(kq.getDepartment())
	// || !kqPersons.get(j).getName().equals(kq.getName())) {
	// KQPerson tempKQPerson = new KQPerson();
	// tempKQPerson.setDepartment(kq.getDepartment());
	// tempKQPerson.setName(kq.getName());
	// List<KQ> tempKQs = new ArrayList<KQ>();
	// tempKQs.add(kq);
	// tempKQPerson.setKqs(tempKQs);
	// kqPersons.add(tempKQPerson);
	// }
	// } else {
	// KQPerson kqPerson = new KQPerson();
	// kqPerson.setDepartment(kq.getDepartment());
	// kqPerson.setName(kq.getName());
	// List<KQ> tempKQs = new ArrayList<KQ>();
	// tempKQs.add(kq);
	// kqPerson.setKqs(tempKQs);
	// kqPersons.add(kqPerson);
	// }
	// }
	//
	// for (KQPerson kqPerson : kqPersons) {
	// String fileName = kqPerson.getDepartment() + "-"
	// + kqPerson.getName();
	// // System.out.println(fileName);
	// Workbook wb = Tools.openPerson(fileName, Var.XLS);
	// for (KQ kq : kqPerson.getKqs()) {
	// Sheet sheet = wb.getSheetAt(0);
	// String date = kq.getDate();
	// for (int i = 4; i <= sheet.getLastRowNum(); i++) {
	// if (sheet.getRow(i).getCell(0).getStringCellValue()
	// .equals(date)) {
	// Tools.writeString(wb, i, 3, kq.getStart(), true);
	// Tools.writeString(wb, i, 4, kq.getEnd(), true);
	// break;
	// }
	// }
	// }
	// Tools.savePerson(fileName, Var.XLS, wb);
	// }
	// }

	/*
	 * public static void setOut() throws IOException, ParseException {
	 * System.out.println("******setOut******");
	 * 
	 * List<List<Object>> outs = Tools.readAll(Var.getOut());
	 * 
	 * List<Out> outList = new ArrayList<Out>(); if (outs == null) { return; }
	 * for (int i = 3; i < outs.size(); i++) { Out out = new Out();
	 * out.setDate(((String) outs.get(i).get(3)).substring(5));
	 * out.setDepartment((String) outs.get(i).get(2)); out.setName((String)
	 * outs.get(i).get(1));
	 * 
	 * String start = ((String) outs.get(i).get(4)).substring(11); String end =
	 * ((String) outs.get(i).get(5)).substring(11);
	 * 
	 * SimpleDateFormat sdf = new SimpleDateFormat("HH:mm");
	 * 
	 * Date start1 = sdf.parse(start); Date end1 = sdf.parse(end); int time = 0;
	 * if (start1.before(sdf.parse("09:01"))) { time += 1; } if
	 * (end1.after(sdf.parse("17:29"))) { time += 2; } out.setTime(time);
	 * 
	 * outList.add(out); }
	 * 
	 * outPersons = new ArrayList<OutPerson>();
	 * 
	 * for (Out out : outList) { if (outPersons.size() > 0) { int j = 0; for
	 * (int i = 0; i < outPersons.size(); i++) { if
	 * (outPersons.get(i).getDepartment() .equals(out.getDepartment()) &&
	 * outPersons.get(i).getName() .equals(out.getName())) { List<Out> tempOuts
	 * = outPersons.get(i).getOuts(); tempOuts.add(out);
	 * outPersons.get(i).setOuts(tempOuts); j = i; break; } } if
	 * (!outPersons.get(j).getDepartment() .equals(out.getDepartment()) ||
	 * !outPersons.get(j).getName().equals(out.getName())) { OutPerson
	 * tempOutPerson = new OutPerson();
	 * tempOutPerson.setDepartment(out.getDepartment());
	 * tempOutPerson.setName(out.getName()); List<Out> tempOuts = new
	 * ArrayList<Out>(); tempOuts.add(out); tempOutPerson.setOuts(tempOuts);
	 * outPersons.add(tempOutPerson); } } else { OutPerson outPerson = new
	 * OutPerson(); outPerson.setDepartment(out.getDepartment());
	 * outPerson.setName(out.getName()); List<Out> tempOuts = new
	 * ArrayList<Out>(); tempOuts.add(out); outPerson.setOuts(tempOuts);
	 * outPersons.add(outPerson); } }
	 * 
	 * for (OutPerson outPerson : outPersons) { String fileName =
	 * outPerson.getDepartment() + "-" + outPerson.getName(); //
	 * System.out.println(fileName); Workbook wb = Tools.openPerson(fileName,
	 * Var.XLS); for (Out out : outPerson.getOuts()) { //
	 * System.out.println(outPerson.getOuts().size()); Sheet sheet =
	 * wb.getSheetAt(0); String date = out.getDate(); for (int i = 4; i <
	 * sheet.getLastRowNum(); i++) { if
	 * (sheet.getRow(i).getCell(0).getStringCellValue() .equals(date)) { switch
	 * (out.getTime()) { case 1: if
	 * (sheet.getRow(i).getCell(3).getStringCellValue() .contains("")) {
	 * Tools.writeString(wb, i, 3, "外出", true); } break; case 2: if
	 * (sheet.getRow(i).getCell(4).getStringCellValue() .contains("")) {
	 * Tools.writeString(wb, i, 4, "外出", true); } break; case 3: if
	 * (sheet.getRow(i).getCell(3).getStringCellValue() .contains("")) {
	 * Tools.writeString(wb, i, 3, "外出", true); } if
	 * (sheet.getRow(i).getCell(4).getStringCellValue() .contains("")) {
	 * Tools.writeString(wb, i, 4, "外出", true); } break; default: break; }
	 * break; } } } Tools.savePerson(fileName, Var.XLS, wb); } }
	 */

	// public static void setHoliday() throws IOException {
	// System.out.println("******setHoliday******");
	//
	// List<List<Object>> holidays = Tools.readAll(Var.getHoliday());
	//
	// List<Holiday> hols = new ArrayList<Holiday>();
	// if (holidays == null) {
	// return;
	// }
	// for (int i = 3; i < holidays.size(); i++) {
	// Holiday hol = new Holiday();
	// hol.setDate(((String) holidays.get(i).get(3)).substring(5, 10));
	// hol.setDepartment((String) holidays.get(i).get(2));
	// hol.setName((String) holidays.get(i).get(1));
	// hol.setStart(((String) holidays.get(i).get(3)).substring(5));
	// hol.setEnd(((String) holidays.get(i).get(4)).substring(5));
	// hol.setType((String) holidays.get(i).get(5));
	// double hours = ((Double) holidays.get(i).get(6) * 8 + (Double) holidays
	// .get(i).get(7));
	// hol.setHours(hours);
	// hols.add(hol);
	// }
	//
	// holidayPersons = new ArrayList<HolidayPerson>();
	//
	// for (Holiday holiday : hols) {
	// if (holidayPersons.size() > 0) {
	// int j = 0;
	// for (int i = 0; i < holidayPersons.size(); i++) {
	// if (holidayPersons.get(i).getName()
	// .equals(holiday.getName())
	// && holidayPersons.get(i).getDepartment()
	// .equals(holiday.getDepartment())) {
	// List<Holiday> tempholidays = holidayPersons.get(i)
	// .getHolidays();
	// tempholidays.add(holiday);
	// holidayPersons.get(i).setHolidays(tempholidays);
	// if (holiday.getType().equals("年假")) {
	// holidayPersons.get(i).setNianjia(
	// holidayPersons.get(i).getNianjia()
	// + holiday.getHours());
	// } else if (holiday.getType().equals("事假")) {
	// holidayPersons.get(i).setShijia(
	// holidayPersons.get(i).getShijia()
	// + holiday.getHours());
	// } else if (holiday.getType().equals("病假")) {
	// holidayPersons.get(i).setBingjia(
	// holidayPersons.get(i).getBingjia()
	// + holiday.getHours());
	// } else if (holiday.getType().equals("调休")) {
	// holidayPersons.get(i).setTiaoxiu(
	// holidayPersons.get(i).getTiaoxiu()
	// + holiday.getHours());
	// } else {
	// holidayPersons.get(i).setQita(
	// holidayPersons.get(i).getQita()
	// + holiday.getHours());
	// }
	// j = i;
	// break;
	// }
	// }
	// if (!holidayPersons.get(j).getName().equals(holiday.getName())
	// || !holidayPersons.get(j).getDepartment()
	// .equals(holiday.getDepartment())) {
	// HolidayPerson tempholidayPerson = new HolidayPerson();
	// tempholidayPerson.setName(holiday.getName());
	// tempholidayPerson.setDepartment(holiday.getDepartment());
	// List<Holiday> tempholidays = new ArrayList<Holiday>();
	// tempholidays.add(holiday);
	// tempholidayPerson.setHolidays(tempholidays);
	// if (holiday.getType().equals("年假")) {
	// tempholidayPerson.setNianjia(holiday.getHours());
	// } else if (holiday.getType().equals("事假")) {
	// tempholidayPerson.setShijia(holiday.getHours());
	// } else if (holiday.getType().equals("病假")) {
	// tempholidayPerson.setBingjia(holiday.getHours());
	// } else if (holiday.getType().equals("调休")) {
	// tempholidayPerson.setTiaoxiu(holiday.getHours());
	// } else {
	// tempholidayPerson.setQita(holiday.getHours());
	// }
	// holidayPersons.add(tempholidayPerson);
	// }
	// } else {
	// HolidayPerson holidayPerson = new HolidayPerson();
	// holidayPerson.setName(holiday.getName());
	// holidayPerson.setDepartment(holiday.getDepartment());
	// List<Holiday> tempholidays = new ArrayList<Holiday>();
	// tempholidays.add(holiday);
	// holidayPerson.setHolidays(tempholidays);
	// if (holiday.getType().equals("年假")) {
	// holidayPerson.setNianjia(holiday.getHours());
	// } else if (holiday.getType().equals("事假")) {
	// holidayPerson.setShijia(holiday.getHours());
	// } else if (holiday.getType().equals("病假")) {
	// holidayPerson.setBingjia(holiday.getHours());
	// } else if (holiday.getType().equals("调休")) {
	// holidayPerson.setTiaoxiu(holiday.getHours());
	// } else {
	// holidayPerson.setQita(holiday.getHours());
	// }
	// holidayPersons.add(holidayPerson);
	// }
	// }
	//
	// for (HolidayPerson holidayPerson : holidayPersons) {
	// String fileName = holidayPerson.getDepartment() + "-"
	// + holidayPerson.getName();
	// System.out.println(fileName);
	// Workbook wb = Tools.openPerson(fileName, Var.XLS);
	// Sheet sheet1 = wb.getSheetAt(0);
	// // for (int i = 0; i <= sheet1.getLastRowNum(); i++) {
	// // if (sheet1.getRow(i).getCell(0).getStringCellValue()
	// // .equals("年假：")) {
	// // Tools.writeDouble(wb, i, 1, holidayPerson.getNianjia(),
	// // false);
	// // Tools.writeDouble(wb, i, 4, holidayPerson.getBingjia(),
	// // false);
	// // Tools.writeDouble(wb, i, 7, holidayPerson.getShijia(),
	// // false);
	// // Tools.writeDouble(wb, i, 10, holidayPerson.getTiaoxiu(),
	// // false);
	// // Tools.writeDouble(wb, i, 13, holidayPerson.getQita(), false);
	// // break;
	// // }
	// // }
	// for (Holiday holiday : holidayPerson.getHolidays()) {
	// // Workbook wb = Tools.open(outPath, fileName, XLS);
	// Sheet sheet = wb.getSheetAt(0);
	// String date = holiday.getDate();
	// for (int i = 4; i <= sheet.getLastRowNum(); i++) {
	// if (sheet.getRow(i).getCell(0).getStringCellValue()
	// .equals(date)) {
	// if (sheet.getRow(i).getCell(10).getStringCellValue()
	// .equals("")
	// || sheet.getRow(i).getCell(10) == null) {
	// Tools.writeString(wb, i, 10, holiday.getStart(),
	// true);
	// Tools.writeString(wb, i, 11, holiday.getEnd(), true);
	// Tools.writeDouble(wb, i, 12, holiday.getHours(),
	// true);
	// Tools.writeString(wb, i, 13, holiday.getType(),
	// true);
	// break;
	// } else if (!sheet.getRow(i + 1).getCell(0)
	// .getStringCellValue().equals(date)) {
	// // System.out.println("" + i);
	// Tools.shift(wb, i + 1);
	// Tools.writeString(wb, i + 1, 10,
	// holiday.getStart(), true);
	// Tools.writeString(wb, i + 1, 11, holiday.getEnd(),
	// true);
	// Tools.writeDouble(wb, i + 1, 12,
	// holiday.getHours(), true);
	// Tools.writeString(wb, i + 1, 13, holiday.getType(),
	// true);
	// break;
	// }
	// }
	// }
	// }
	// Tools.savePerson(fileName, Var.XLS, wb);
	// }
	// }

	// public static void setAdd() throws IOException, ParseException {
	// System.out.println("******setAdd******");
	//
	// List<List<Object>> add1s = Tools.readAll(Var.getAdd());
	// // List<List<Object>> add2s = Tools.readAll(Var.getAdd2());
	// // List<List<Object>> add2ses = Tools.readAll(Var.getAdd());
	//
	// List<Add> adds = new ArrayList<Add>();
	//
	// // if (add2s != null) {
	// // for (int i = 3; i < add2s.size(); i++) {
	// // Add add = new Add();
	// // add.setDate(((String) add2s.get(i).get(4)).substring(5, 10));
	// // add.setDepartment((String) add2s.get(i).get(2));
	// // add.setName((String) add2s.get(i).get(1));
	// // add.setStart(((String) add2s.get(i).get(4)).substring(5));
	// // add.setEnd(((String) add2s.get(i).get(5)).substring(5));
	// // add.setSite((String) add2s.get(i).get(7));
	// // add.setApply("是");
	// // double hours = (Double) add2s.get(i).get(6) * 24;
	// //
	// // BigDecimal b = new BigDecimal(hours);
	// // add.setHours(b.setScale(1, BigDecimal.ROUND_HALF_UP)
	// // .doubleValue());
	// //
	// // adds.add(add);
	// // }
	// //
	// // if (add2ses != null) {
	// // for (int i = 3; i < add2ses.size(); i++) {
	// // for (int j = 0; j < adds.size(); j++) {
	// // if (adds.get(j).getName()
	// // .equals((String) add2ses.get(i).get(1))
	// // && adds.get(j).getDepartment()
	// // .equals((String) add2ses.get(i).get(2))
	// // && adds.get(j)
	// // .getStart()
	// // .equals(((String) add2ses.get(i).get(4))
	// // .substring(5))
	// // && adds.get(j)
	// // .getEnd()
	// // .equals(((String) add2ses.get(i).get(5))
	// // .substring(5))) {
	// // adds.get(j).setApply("否");
	// // break;
	// // }
	// // }
	// // }
	// // }
	// // }
	//
	// if (add1s != null) {
	// for (int i = 3; i < add1s.size(); i++) {
	// Add add = new Add();
	// add.setDate(((String) add1s.get(i).get(4)).substring(5, 10));
	// add.setDepartment((String) add1s.get(i).get(2));
	// add.setName((String) add1s.get(i).get(1));
	// add.setStart(((String) add1s.get(i).get(4)).substring(5));
	// add.setEnd(((String) add1s.get(i).get(5)).substring(5));
	// add.setSite((String) add1s.get(i).get(7));
	// add.setApply("");
	// SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm");
	// Date start = sdf.parse((String) add1s.get(i).get(4));
	// Date end = sdf.parse((String) add1s.get(i).get(5));
	// long diff = end.getTime() - start.getTime();
	// double hours = (double) diff / 60 / 60 / 1000;
	// BigDecimal b = new BigDecimal(hours);
	// add.setHours(b.setScale(1, BigDecimal.ROUND_HALF_UP)
	// .doubleValue());
	//
	// adds.add(add);
	// }
	// }
	// addPersons = new ArrayList<AddPerson>();
	//
	// for (Add add : adds) {
	// if (addPersons.size() > 0) {
	// int j = 0;
	// for (int i = 0; i < addPersons.size(); i++) {
	// if (addPersons.get(i).getDepartment()
	// .equals(add.getDepartment())
	// && addPersons.get(i).getName()
	// .equals(add.getName())) {
	// List<Add> tempAdds = addPersons.get(i).getAdds();
	// tempAdds.add(add);
	// addPersons.get(i).setAdds(tempAdds);
	// j = i;
	// break;
	// }
	// }
	// if (!addPersons.get(j).getDepartment()
	// .equals(add.getDepartment())
	// || !addPersons.get(j).getName().equals(add.getName())) {
	// AddPerson tempAddPerson = new AddPerson();
	// tempAddPerson.setDepartment(add.getDepartment());
	// tempAddPerson.setName(add.getName());
	// List<Add> tempAdds = new ArrayList<Add>();
	// tempAdds.add(add);
	// tempAddPerson.setAdds(tempAdds);
	// addPersons.add(tempAddPerson);
	// }
	// } else {
	// AddPerson tempAddPerson = new AddPerson();
	// tempAddPerson.setDepartment(add.getDepartment());
	// tempAddPerson.setName(add.getName());
	// List<Add> tempAdds = new ArrayList<Add>();
	// tempAdds.add(add);
	// tempAddPerson.setAdds(tempAdds);
	// addPersons.add(tempAddPerson);
	// }
	// }
	//
	// for (AddPerson addPerson : addPersons) {
	// String fileName = addPerson.getDepartment() + "-"
	// + addPerson.getName();
	// // System.out.println(fileName);
	// Workbook wb = Tools.openPerson(fileName, Var.XLS);
	// Sheet sheet = wb.getSheetAt(0);
	// for (Add add : addPerson.getAdds()) {
	// String date = add.getDate();
	// for (int i = 4; i < sheet.getLastRowNum(); i++) {
	// if (sheet.getRow(i).getCell(0).getStringCellValue()
	// .equals(date)) {
	// if (sheet.getRow(i).getCell(6).getStringCellValue()
	// .equals("")) {
	// Tools.writeString(wb, i, 5, add.getSite(), true);
	// Tools.writeString(wb, i, 6, add.getStart(), true);
	// Tools.writeString(wb, i, 7, add.getEnd(), true);
	// Tools.writeDouble(wb, i, 8, add.getHours(), true);
	// Tools.writeString(wb, i, 9, add.getApply(), true);
	// break;
	// } else if (!sheet.getRow(i + 1).getCell(0)
	// .getStringCellValue().equals(date)) {
	// // System.out.println("" + i);
	// Tools.shift(wb, i + 1);
	// Tools.writeString(wb, i + 1, 5, add.getSite(), true);
	// Tools.writeString(wb, i + 1, 6, add.getStart(),
	// true);
	// Tools.writeString(wb, i + 1, 7, add.getEnd(), true);
	// Tools.writeDouble(wb, i + 1, 8, add.getHours(),
	// true);
	// Tools.writeString(wb, i + 1, 9, add.getApply(),
	// true);
	// break;
	// }
	// }
	// }
	// }
	// Tools.savePerson(fileName, Var.XLS, wb);
	// }
	// }

	// public static void setAdd() throws IOException, ParseException {
	// System.out.println("******setAdd******");
	//
	// List<List<Object>> add1s = Tools.readAll(Var.getAdd1());
	// List<List<Object>> add2s = Tools.readAll(Var.getAdd2());
	// List<List<Object>> add2ses = Tools.readAll(Var.getAdd2se());
	//
	// List<Add> adds = new ArrayList<Add>();
	//
	// if (add2s != null) {
	// for (int i = 3; i < add2s.size(); i++) {
	// Add add = new Add();
	// add.setDate(((String) add2s.get(i).get(4)).substring(5, 10));
	// add.setDepartment((String) add2s.get(i).get(2));
	// add.setName((String) add2s.get(i).get(1));
	// add.setStart(((String) add2s.get(i).get(4)).substring(5));
	// add.setEnd(((String) add2s.get(i).get(5)).substring(5));
	// add.setSite((String) add2s.get(i).get(7));
	// add.setApply("是");
	// double hours = (Double) add2s.get(i).get(6) * 24;
	//
	// BigDecimal b = new BigDecimal(hours);
	// add.setHours(b.setScale(1, BigDecimal.ROUND_HALF_UP)
	// .doubleValue());
	//
	// adds.add(add);
	// }
	//
	// if (add2ses != null) {
	// for (int i = 3; i < add2ses.size(); i++) {
	// for (int j = 0; j < adds.size(); j++) {
	// if (adds.get(j).getName()
	// .equals((String) add2ses.get(i).get(1))
	// && adds.get(j).getDepartment()
	// .equals((String) add2ses.get(i).get(2))
	// && adds.get(j)
	// .getStart()
	// .equals(((String) add2ses.get(i).get(4))
	// .substring(5))
	// && adds.get(j)
	// .getEnd()
	// .equals(((String) add2ses.get(i).get(5))
	// .substring(5))) {
	// adds.get(j).setApply("否");
	// break;
	// }
	// }
	// }
	// }
	// }
	//
	// if (add1s != null) {
	// for (int i = 3; i < add1s.size(); i++) {
	// Add add = new Add();
	// add.setDate(((String) add1s.get(i).get(3)).substring(5, 10));
	// add.setDepartment((String) add1s.get(i).get(2));
	// add.setName((String) add1s.get(i).get(1));
	// add.setStart(((String) add1s.get(i).get(3)).substring(5));
	// add.setEnd(((String) add1s.get(i).get(4)).substring(5));
	// add.setSite((String) add1s.get(i).get(7));
	// add.setApply("是");
	// SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm");
	// Date start = sdf.parse((String) add1s.get(i).get(3));
	// Date end = sdf.parse((String) add1s.get(i).get(4));
	// long diff = end.getTime() - start.getTime();
	// double hours = (double) diff / 60 / 60 / 1000;
	// BigDecimal b = new BigDecimal(hours);
	// add.setHours(b.setScale(1, BigDecimal.ROUND_HALF_UP)
	// .doubleValue());
	//
	// adds.add(add);
	// }
	// }
	// addPersons = new ArrayList<AddPerson>();
	//
	// for (Add add : adds) {
	// if (addPersons.size() > 0) {
	// int j = 0;
	// for (int i = 0; i < addPersons.size(); i++) {
	// if (addPersons.get(i).getDepartment()
	// .equals(add.getDepartment())
	// && addPersons.get(i).getName()
	// .equals(add.getName())) {
	// List<Add> tempAdds = addPersons.get(i).getAdds();
	// tempAdds.add(add);
	// addPersons.get(i).setAdds(tempAdds);
	// j = i;
	// break;
	// }
	// }
	// if (!addPersons.get(j).getDepartment()
	// .equals(add.getDepartment())
	// || !addPersons.get(j).getName().equals(add.getName())) {
	// AddPerson tempAddPerson = new AddPerson();
	// tempAddPerson.setDepartment(add.getDepartment());
	// tempAddPerson.setName(add.getName());
	// List<Add> tempAdds = new ArrayList<Add>();
	// tempAdds.add(add);
	// tempAddPerson.setAdds(tempAdds);
	// addPersons.add(tempAddPerson);
	// }
	// } else {
	// AddPerson tempAddPerson = new AddPerson();
	// tempAddPerson.setDepartment(add.getDepartment());
	// tempAddPerson.setName(add.getName());
	// List<Add> tempAdds = new ArrayList<Add>();
	// tempAdds.add(add);
	// tempAddPerson.setAdds(tempAdds);
	// addPersons.add(tempAddPerson);
	// }
	// }
	//
	// for (AddPerson addPerson : addPersons) {
	// String fileName = addPerson.getDepartment() + "-"
	// + addPerson.getName();
	// // System.out.println(fileName);
	// Workbook wb = Tools.openPerson(fileName, Var.XLS);
	// Sheet sheet = wb.getSheetAt(0);
	// for (Add add : addPerson.getAdds()) {
	// String date = add.getDate();
	// for (int i = 4; i < sheet.getLastRowNum(); i++) {
	// if (sheet.getRow(i).getCell(0).getStringCellValue()
	// .equals(date)) {
	// if (sheet.getRow(i).getCell(6).getStringCellValue() == "") {
	// Tools.writeString(wb, i, 5, add.getSite(), true);
	// Tools.writeString(wb, i, 6, add.getStart(), true);
	// Tools.writeString(wb, i, 7, add.getEnd(), true);
	// Tools.writeDouble(wb, i, 8, add.getHours(), true);
	// Tools.writeString(wb, i, 9, add.getApply(), true);
	// break;
	// } else if (!sheet.getRow(i + 1).getCell(0)
	// .getStringCellValue().equals(date)) {
	// // System.out.println("" + i);
	// Tools.shift(wb, i + 1);
	// Tools.writeString(wb, i + 1, 5, add.getSite(), true);
	// Tools.writeString(wb, i + 1, 6, add.getStart(),
	// true);
	// Tools.writeString(wb, i + 1, 7, add.getEnd(), true);
	// Tools.writeDouble(wb, i + 1, 8, add.getHours(),
	// true);
	// Tools.writeString(wb, i + 1, 9, add.getApply(),
	// true);
	// break;
	// }
	// }
	// }
	// }
	// Tools.savePerson(fileName, Var.XLS, wb);
	// }
	// }

	public static void fixPB() throws IOException {
		System.out.println("******fixPB******");

		for (PBKQPerson pbPerson : pbkqPersons) {
			String fileName = pbPerson.getDepartment() + "-"
					+ pbPerson.getName();
			// System.out.println(fileName);
			Workbook wb = Tools.openPerson(fileName, Var.XLS);
			if (wb != null) {
				Sheet sheet = wb.getSheetAt(0);

				// double weekday = 0;
				// double weekend = 0;
				// double holiday = 0;
				// for (int i = 4; i < sheet.getLastRowNum(); i++) {
				// Row row = sheet.getRow(i);
				//
				// if (row.getCell(8) != null) {
				// // System.out.println(i);
				// if (row.getCell(2) != null
				// && (row.getCell(2).getStringCellValue()
				// .contains("节") || row.getCell(2)
				// .getStringCellValue().contains("元旦"))) {
				// holiday = holiday
				// + row.getCell(8).getNumericCellValue();
				// } else if (row.getCell(2) != null
				// && (row.getCell(2).getStringCellValue()
				// .equals("休息") || row.getCell(2)
				// .getStringCellValue().contains("加班"))) {
				// weekend = weekend
				// + row.getCell(8).getNumericCellValue();
				// } else {
				// weekday = weekday
				// + row.getCell(8).getNumericCellValue();
				// }
				// }
				// }

				for (int i = 4; i < sheet.getLastRowNum(); i++) {
					Row row = sheet.getRow(i);

					if (row.getCell(2) != null) {
						if ((row.getCell(2).getStringCellValue().contains("中班") || row
								.getCell(2).getStringCellValue().contains("晚班"))
								&& !row.getCell(10).getStringCellValue()
										.equals("")
								&& !row.getCell(0)
										.getStringCellValue()
										.equals(sheet.getRow(i - 1).getCell(0)
												.getStringCellValue())) {
							int temp = 0;
							temp = (int) (row.getCell(12).getNumericCellValue() / 8);
							for (int j = 0; j < temp; j++) {
								if (sheet.getRow(i + j).getCell(2)
										.getStringCellValue().contains("中班")) {
									pbPerson.setZhongban(pbPerson.getZhongban() - 1);
									System.out.println(pbPerson.getName()
											+ "zhong-" + 1);
								} else if (sheet.getRow(i + j).getCell(2)
										.getStringCellValue().contains("晚班")) {
									pbPerson.setYeban(pbPerson.getYeban() - 1);
									System.out.println(pbPerson.getName()
											+ "ye-" + 1);
								}
							}

						}

						if (row.getCell(6).getStringCellValue()
								.contains("21:00")
								&& row.getCell(7).getStringCellValue()
										.contains("09:00")
								&& !row.getCell(0)
										.getStringCellValue()
										.equals(sheet.getRow(i - 1).getCell(0)
												.getStringCellValue())) {
							pbPerson.setYeban(pbPerson.getYeban() + 1);
							System.out.println(pbPerson.getName() + "ye+");
						}
					}
				}

				for (int i = 36; i <= sheet.getLastRowNum(); i++) {
					if (sheet.getRow(i).getCell(0).getStringCellValue()
							.equals("中班数：")) {
						Tools.writeString(wb, i, 1,
								"" + pbPerson.getZhongban(), 3);
						Tools.writeString(wb, i, 4, "" + pbPerson.getYeban(), 3);
						// Tools.writeDouble(wb, i, 7, weekday, false);
						// Tools.writeDouble(wb, i, 10, weekend, false);
						// Tools.writeDouble(wb, i, 13, holiday, false);
					}
				}

				Tools.savePerson(fileName, Var.XLS, wb);
			}
		}
	}

	public static void setRowHeight() throws IOException {
		System.out.println("******setRowHeight******");

		for (Person person : persons) {
			String fileName = person.getDepartment() + "-" + person.getName();
			Workbook wb = Tools.openPerson(fileName, Var.XLS);
			Sheet sheet = wb.getSheetAt(0);
			for (int i = 1; i < sheet.getLastRowNum(); i++) {
				Row row = sheet.getRow(i);

				if (row.getCell(0) == null) {
					break;
				}

				Cell cell = row.getCell(5);
				if (cell != null && !cell.toString().equals("")) {
					int rows = cell.toString().length() / 6;
					row.setHeight((short) (row.getHeight() * (rows + 1)));
				}
			}
			Tools.savePerson(fileName, Var.XLS, wb);
		}
	}

	@SuppressWarnings("deprecation")
	public static void merge() throws IOException {
		System.out.println("******merge******");

		for (Person person : persons) {
			String fileName = person.getDepartment() + "-" + person.getName();
			// System.out.println(fileName);
			Workbook wb = Tools.openPerson(fileName, Var.XLS);
			Sheet sheet = wb.getSheetAt(0);
			boolean flag = false;
			for (int i = 4; i < sheet.getLastRowNum(); i++) {
				Row row = sheet.getRow(i);
				Cell cell = row.getCell(0);

				int k = 0;
				while (sheet.getRow(i + k + 1).getCell(0).getStringCellValue()
						.equals(cell.getStringCellValue())
						&& !cell.getStringCellValue().equals("")
						&& cell != null) {
					k++;
				}
				if (k > 0) {
					flag = true;
					CellStyle style = wb.createCellStyle();
					style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
					style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
					style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
					style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
					style.setBorderRight(HSSFCellStyle.BORDER_THIN);
					style.setBorderTop(HSSFCellStyle.BORDER_THIN);

					sheet.addMergedRegion(new CellRangeAddress(i, i + k, 0, 0));
					sheet.getRow(i).getCell(0).setCellStyle(style);
					sheet.addMergedRegion(new CellRangeAddress(i, i + k, 1, 1));
					sheet.getRow(i).getCell(1).setCellStyle(style);
					sheet.addMergedRegion(new CellRangeAddress(i, i + k, 2, 2));
					sheet.getRow(i).getCell(2).setCellStyle(style);
					sheet.addMergedRegion(new CellRangeAddress(i, i + k, 3, 3));
					sheet.getRow(i).getCell(3).setCellStyle(style);
					sheet.addMergedRegion(new CellRangeAddress(i, i + k, 4, 4));
					sheet.getRow(i).getCell(4).setCellStyle(style);

					System.out.println(fileName + i + "," + (i + k));
					i += k;
				}

				if (flag) {
					if (cell.toString().equals("合计")) {
						sheet.addMergedRegion(new CellRangeAddress(i, i, 0, 4));
						sheet.addMergedRegion(new CellRangeAddress(i, i, 6, 7));
						sheet.addMergedRegion(new CellRangeAddress(i, i, 10, 11));
					}
				}

				if (cell.toString().contains("备注：")) {
					if (!Tools.isMerged(sheet, i, 0)) {
						sheet.addMergedRegion(new CellRangeAddress(i, i, 0, 14));
					}
				}
			}
			Tools.savePerson(fileName, Var.XLS, wb);
		}
	}

}
