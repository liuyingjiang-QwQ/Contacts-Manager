#include <iostream>
#include <Windows.h>
#include <string.h>
#include <conio.h>
#include <xlnt/xlnt.hpp>
using namespace std;
int line_number;    //xlsx表格中的行号
int contact_counter;    //list表格时联系人的排序
string search_name;    //你输入的需要寻找的名字
string order;
string write;    //添加联系人时写入的内容
string cell;    //单元格
void list_contacts() {
	SetConsoleOutputCP(65001);
	SetConsoleCP(65001);
	HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);    //使用UTF-8编码，否则的话无法正常读写中文
	xlnt::workbook wb;
	wb.load("Memo.xlsx");
	xlnt::worksheet ws = wb.sheet_by_title("Sheet1");
	contact_counter = 1;
	xlnt::row_t first_row = ws.lowest_row();
	xlnt::row_t last_row = ws.highest_row();
	if (last_row <= 1) {
		cout << "没有联系人";
		cout << "\n按下任意按键继续...";
		while (true) {
			if (_kbhit()) {
				break;
			}
			Sleep(10);
		}
	}
	else {
		for (xlnt::row_t r = first_row + 1; r <= last_row; r++) {
			cell = string("A") + to_string(r);
			if (ws.cell(cell).value<string>() != "") {
				SetConsoleTextAttribute(hConsole, 2);
				cout << "\n联系人 " << contact_counter << endl;
				SetConsoleTextAttribute(hConsole, 7);
				cout << "姓名：    " << ws.cell(cell).value<string>() << endl;
				cell = string("B") + to_string(r);
				cout << "生日：    " << ws.cell(cell).value<string>() << endl;
				cell = string("C") + to_string(r);
				cout << "住址：    " << ws.cell(cell).value<string>() << endl;
				cell = string("D") + to_string(r);
				cout << "联系方式：" << ws.cell(cell).value<string>() << endl;
				cell = string("E") + to_string(r);
				cout << "备注：    " << ws.cell(cell).value<string>() << endl;
				contact_counter++;
			}
		}
		cout << "\n以上是所有联系人的信息\n";
	}
}
int main() {
	SetConsoleOutputCP(65001);
	SetConsoleCP(65001);
	HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
	try {    //尝试打开Memo.xlsx表格，如果打开失败（可能是因为文件不存在），则创建新文件
		xlnt::workbook wb;
		wb.load("Memo.xlsx");
		xlnt::worksheet ws = wb.sheet_by_title("Sheet1");
	}
	catch (const std::exception& e) {
		xlnt::workbook wb;
		xlnt::worksheet ws = wb.sheet_by_title("Sheet1");
		ws.cell("A1").value("姓名");
		ws.cell("B1").value("生日");
		ws.cell("C1").value("住址");
		ws.cell("D1").value("联系方式");
		ws.cell("E1").value("备注");
		wb.save("Memo.xlsx");
	}
	xlnt::workbook wb;
	wb.load("Memo.xlsx");
	xlnt::worksheet ws = wb.sheet_by_title("Sheet1");
	while (true) {
	Da_Capo:    //或许这是我作为音乐爱好者埋藏的一些彩蛋？（在音乐中，Da Capo意思是回到开头，从头再奏）
		system("cls||clear");
		cout << "╔══════════════════════╗\n"
			<< "║      亲朋备忘录      ║\n"
			<< "╠══════════════════════╣\n"
			<< "║ 1. 写入内容          ║\n"
			<< "║ 2. 查询联系人        ║\n"
			<< "║ 3. 删除联系人        ║\n"
			<< "║ 4. 退出              ║\n"
			<< "║ 输入\"list\"以列出表格 ║\n"
			<< "╚══════════════════════╝\n"
			<< "输入选项：";
		getline(cin, order);
		if (order == "1" || order == "写入内容" || order == "写入" || order == "写入联系人") {
			line_number = 2;
			while (true) {
				cell = string("A") + to_string(line_number);
				if (ws.cell(cell).value<string>() == "") {    //因为在删除联系人时，我并没有设计让下方的数据自动调到上方的空行，为了这些空行不被浪费，就用空行来写入新联系人吧
					system("cls||clear");
					cout << "请输入他/她/它的姓名：";
					getline(cin, write);
					xlnt::row_t first_row = ws.lowest_row();
					xlnt::row_t last_row = ws.highest_row();
					for (xlnt::row_t r = first_row; r <= last_row; r++) {
						cell = string("A") + to_string(r);
						if (ws.cell(cell).value<string>() == write) {
							cout << "\n输入的联系人姓名已存在于表格中！";
							cout << "\n按下任意按键继续...";
							while (true) {
								if (_kbhit()) {
									break;
								}
								Sleep(10);
							}
							goto Da_Capo;    //这里直接回到开头，也就是说下面的程序就不会执行，只有上方的if不成立（不运行这段）才会运行下面的程序。虽然逻辑乱了一点，但是能够达成目的就行
						}
					}
					cell = string("A") + to_string(line_number);
					ws.cell(cell).value(write);
					cell = string("B") + to_string(line_number);
					cout << "请输入他/她/它的生日：";
					getline(cin, write);
					ws.cell(cell).value(write);
					cell = string("C") + to_string(line_number);
					cout << "请输入他/她/它的住址：";
					getline(cin, write);
					ws.cell(cell).value(write);
					cell = string("D") + to_string(line_number);
					cout << "请输入他/她/它的联系方式：";
					getline(cin, write);
					ws.cell(cell).value(write);
					cell = string("E") + to_string(line_number);
					cout << "请输入你对他/她/它的备注：";
					getline(cin, write);
					ws.cell(cell).value(write);
					try {
						wb.save("Memo.xlsx");
					}
					catch (const xlnt::exception& e) {
						std::cerr << "Error saving workbook: " << e.what() << std::endl;
					}
					cout << "已将联系人信息存入表格中！\n";
					cout << "按下任意按键继续...";
					while (true) {
						if (_kbhit()) {
							break;
						}
						Sleep(10);
					}
					break;
				}
				line_number++;
			}
		}
		else if (order == "list" || order == "List") {
			list_contacts();
			cout << "按下任意按键继续...";
			while (true) {
				if (_kbhit()) {
					break;
				}
				Sleep(10);
			}
		}
		else if (order == "2" || order == "查询" || order == "查询联系人") {
			system("cls||clear");
			cout << "输入你要寻找的联系人：";
			getline(cin, search_name);
			xlnt::row_t first_row = ws.lowest_row();
			xlnt::row_t last_row = ws.highest_row();
			for (xlnt::row_t r = first_row + 1; r <= last_row; r++) {
				cell = string("A") + to_string(r);
				if (ws.cell(cell).value<string>() == search_name) {
					cout << "\n联系人<" << ws.cell(cell).value<string>() << ">的信息：\n";
					cout << "姓名：    " << ws.cell(cell).value<string>() << endl;
					cell = string("B") + to_string(r);
					cout << "生日：    " << ws.cell(cell).value<string>() << endl;
					cell = string("C") + to_string(r);
					cout << "住址：    " << ws.cell(cell).value<string>() << endl;
					cell = string("D") + to_string(r);
					cout << "联系方式：" << ws.cell(cell).value<string>() << endl;
					cell = string("E") + to_string(r);
					cout << "备注：    " << ws.cell(cell).value<string>() << endl;
					cout << "\n按下任意按键继续...";
					while (true) {
						if (_kbhit()) {
							break;
						}
						Sleep(10);
					}
					goto Da_Capo;
				}
			}
			cout << "未找到此联系人";
			cout << "\n按下任意按键继续...";
			while (true) {
				if (_kbhit()) {
					break;
				}
				Sleep(10);
			}
		}
		else if (order == "3" || order == "删除" || order == "删除联系人") {
			system("cls||clear");
			while (true) {
				cout << "输入你要删除的联系人姓名，或输入list列出他/她/它\n>>>";
				getline(cin, search_name);
				if (search_name == "list" || search_name == "List") {
					list_contacts();
				}
				else {
					xlnt::row_t first_row = ws.lowest_row();
					xlnt::row_t last_row = ws.highest_row();
					for (xlnt::row_t r = first_row + 1; r <= last_row; r++) {
						cell = string("A") + to_string(r);
						if (ws.cell(cell).value<string>() == search_name && ws.cell(cell).value<string>() != "") {
							cout << "\n联系人<" << ws.cell(cell).value<string>() << ">的信息：\n";
							cout << "姓名：    " << ws.cell(cell).value<string>() << endl;
							cell = string("B") + to_string(r);
							cout << "生日：    " << ws.cell(cell).value<string>() << endl;
							cell = string("C") + to_string(r);
							cout << "住址：    " << ws.cell(cell).value<string>() << endl;
							cell = string("D") + to_string(r);
							cout << "联系方式：" << ws.cell(cell).value<string>() << endl;
							cell = string("E") + to_string(r);
							cout << "备注：    " << ws.cell(cell).value<string>() << endl;
							cout << "\n你确定要删除他/她/它吗？\ny.是      n.不是\n>>>";
							getline(cin, order);
							if (order == "y" || order == "Y") {
								try {
									// Clear all fields for this contact
									cell = string("A") + to_string(r);
									ws.cell(cell).value("");
									cell = string("B") + to_string(r);
									ws.cell(cell).value("");
									cell = string("C") + to_string(r);
									ws.cell(cell).value("");
									cell = string("D") + to_string(r);
									ws.cell(cell).value("");
									cell = string("E") + to_string(r);
									ws.cell(cell).value("");
									wb.save("Memo.xlsx");
									cout << "\n已成功删除联系人！";
									cout << "\n按下任意键继续...";
									while (true) {
										if (_kbhit()) {
											break;
										}
										Sleep(10);
									}
									break;
								}
								catch (const xlnt::exception& e) {
									std::cerr << "\n\nError saving workbook: " << e.what() << std::endl;
									cout << "\n\n按下任意键继续...";
									while (true) {
										if (_kbhit()) {
											break;
										}
										Sleep(10);
									}
									break;
								}
							}
							else {
								cout << "已取消删除联系人";
								cout << "\n按下任意键继续...";
								while (true) {
									if (_kbhit()) {
										break;
									}
									Sleep(10);
								}
								break;
							}
							goto Da_Capo;
						}
					}
					cout << "未找到此联系人";
					cout << "\n按下任意按键继续...";
					while (true) {
						if (_kbhit()) {
							break;
						}
						Sleep(10);
					}
					break;
				}
			}
		}
		else if (order == "4" || order == "退出" || order == "Exit" || order == "exit") {
			break;
		}
	}
	try {
		wb.save("Memo.xlsx");
	}
	catch (const xlnt::exception& e) {
		std::cerr << "Error saving workbook: " << e.what() << std::endl;
	}
	return 0;
}