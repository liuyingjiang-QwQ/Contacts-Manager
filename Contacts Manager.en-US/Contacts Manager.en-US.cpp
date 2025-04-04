#include <iostream>
#include <Windows.h>
#include <string.h>
#include <conio.h>
#include <xlnt/xlnt.hpp>
using namespace std;

int line_number;    // Row number in the xlsx sheet
int contact_counter;    // Counter for listing contacts
string search_name;    // Name to search for
string command;
string input_buffer;    // Content to write when adding a contact
string cell;    // Excel cell reference

void list_contacts() {
    SetConsoleOutputCP(65001);
    SetConsoleCP(65001);
    HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);    // Use UTF-8 encoding to properly handle non-ASCII characters
    xlnt::workbook wb;
    wb.load("Memo.xlsx");
    xlnt::worksheet ws = wb.sheet_by_title("Sheet1");
    contact_counter = 1;
    xlnt::row_t first_row = ws.lowest_row();
    xlnt::row_t last_row = ws.highest_row();

    if (last_row <= 1) {
        cout << "No contacts found";
        cout << "\nPress any key to continue...";
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
                SetConsoleTextAttribute(hConsole, 2);  // Green text for contact number
                cout << "\nContact " << contact_counter << endl;
                SetConsoleTextAttribute(hConsole, 7);   // Reset to default color
                cout << "Name:     " << ws.cell(cell).value<string>() << endl;
                cell = string("B") + to_string(r);
                cout << "Birthday: " << ws.cell(cell).value<string>() << endl;
                cell = string("C") + to_string(r);
                cout << "Address:  " << ws.cell(cell).value<string>() << endl;
                cell = string("D") + to_string(r);
                cout << "Phone:    " << ws.cell(cell).value<string>() << endl;
                cell = string("E") + to_string(r);
                cout << "Notes:    " << ws.cell(cell).value<string>() << endl;
                contact_counter++;
            }
        }
        cout << "\nAbove are all contacts in the database\n";
    }
}

int main() {
    SetConsoleOutputCP(65001);
    SetConsoleCP(65001);
    HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);

    // Try to open Memo.xlsx. If it doesn't exist, create a new file with headers
    try {
        xlnt::workbook wb;
        wb.load("Memo.xlsx");
        xlnt::worksheet ws = wb.sheet_by_title("Sheet1");
    }
    catch (const std::exception& e) {
        xlnt::workbook wb;
        xlnt::worksheet ws = wb.sheet_by_title("Sheet1");
        ws.cell("A1").value("Name");
        ws.cell("B1").value("Birthday");
        ws.cell("C1").value("Address");
        ws.cell("D1").value("Phone");
        ws.cell("E1").value("Notes");
        wb.save("Memo.xlsx");
    }

    xlnt::workbook wb;
    wb.load("Memo.xlsx");
    xlnt::worksheet ws = wb.sheet_by_title("Sheet1");

    while (true) {
    Da_Capo:     // Da Capo (musical term meaning "from the beginning") - our restart point
        system("cls||clear");
        cout << "╔════════════════════════╗\n"
            << "║   Contacts Manager     ║\n"
            << "╠════════════════════════╣\n"
            << "║ 1. Add Contact         ║\n"
            << "║ 2. Search Contact      ║\n"
            << "║ 3. Delete Contact      ║\n"
            << "║ 4. Exit                ║\n"
            << "║ Type \"list\" to show all║\n"
            << "╚════════════════════════╝\n"
            << "Enter option: ";
        getline(cin, command);

        if (command == "1") {
            line_number = 2;
            while (true) {
                cell = string("A") + to_string(line_number);
                // Find the first empty row to add new contact
                if (ws.cell(cell).value<string>() == "") {
                    system("cls||clear");
                    cout << "Enter contact name: ";
                    getline(cin, input_buffer);

                    // Check if contact already exists
                    xlnt::row_t first_row = ws.lowest_row();
                    xlnt::row_t last_row = ws.highest_row();
                    for (xlnt::row_t r = first_row; r <= last_row; r++) {
                        cell = string("A") + to_string(r);
                        if (ws.cell(cell).value<string>() == input_buffer) {
                            cout << "\nThis contact already exists!";
                            cout << "\nPress any key to continue...";
                            while (true) {
                                if (_kbhit()) {
                                    break;
                                }
                                Sleep(10);
                            }
                            goto Da_Capo;
                        }
                    }

                    // Add new contact information
                    cell = string("A") + to_string(line_number);
                    ws.cell(cell).value(input_buffer);

                    cell = string("B") + to_string(line_number);
                    cout << "Enter birthday: ";
                    getline(cin, input_buffer);
                    ws.cell(cell).value(input_buffer);

                    cell = string("C") + to_string(line_number);
                    cout << "Enter address: ";
                    getline(cin, input_buffer);
                    ws.cell(cell).value(input_buffer);

                    cell = string("D") + to_string(line_number);
                    cout << "Enter phone number: ";
                    getline(cin, input_buffer);
                    ws.cell(cell).value(input_buffer);

                    cell = string("E") + to_string(line_number);
                    cout << "Enter notes: ";
                    getline(cin, input_buffer);
                    ws.cell(cell).value(input_buffer);

                    try {
                        wb.save("Memo.xlsx");
                    }
                    catch (const xlnt::exception& e) {
                        std::cerr << "Error saving workbook: " << e.what() << std::endl;
                    }

                    cout << "Contact information saved successfully!\n";
                    cout << "Press any key to continue...";
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
        else if (command == "list" || command == "List") {
            list_contacts();
            cout << "Press any key to continue...";
            while (true) {
                if (_kbhit()) {
                    break;
                }
                Sleep(10);
            }
        }
        else if (command == "2") {
            system("cls||clear");
            cout << "Enter contact name to search: ";
            getline(cin, search_name);
            xlnt::row_t first_row = ws.lowest_row();
            xlnt::row_t last_row = ws.highest_row();

            for (xlnt::row_t r = first_row + 1; r <= last_row; r++) {
                cell = string("A") + to_string(r);
                if (ws.cell(cell).value<string>() == search_name) {
                    cout << "\nInformation for <" << ws.cell(cell).value<string>() << ">:\n";
                    cout << "Name:     " << ws.cell(cell).value<string>() << endl;
                    cell = string("B") + to_string(r);
                    cout << "Birthday: " << ws.cell(cell).value<string>() << endl;
                    cell = string("C") + to_string(r);
                    cout << "Address:  " << ws.cell(cell).value<string>() << endl;
                    cell = string("D") + to_string(r);
                    cout << "Phone:    " << ws.cell(cell).value<string>() << endl;
                    cell = string("E") + to_string(r);
                    cout << "Notes:    " << ws.cell(cell).value<string>() << endl;
                    cout << "\nPress any key to continue...";
                    while (true) {
                        if (_kbhit()) {
                            break;
                        }
                        Sleep(10);
                    }
                    goto Da_Capo;
                }
            }
            cout << "Contact not found";
            cout << "\nPress any key to continue...";
            while (true) {
                if (_kbhit()) {
                    break;
                }
                Sleep(10);
            }
        }
        else if (command == "3") {
            system("cls||clear");
            while (true) {
                cout << "Enter contact name to delete, or type 'list' to show all\n>>>";
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
                            cout << "\nInformation for <" << ws.cell(cell).value<string>() << ">:\n";
                            cout << "Name:     " << ws.cell(cell).value<string>() << endl;
                            cell = string("B") + to_string(r);
                            cout << "Birthday: " << ws.cell(cell).value<string>() << endl;
                            cell = string("C") + to_string(r);
                            cout << "Address:  " << ws.cell(cell).value<string>() << endl;
                            cell = string("D") + to_string(r);
                            cout << "Phone:    " << ws.cell(cell).value<string>() << endl;
                            cell = string("E") + to_string(r);
                            cout << "Notes:    " << ws.cell(cell).value<string>() << endl;
                            cout << "\nAre you sure you want to delete this contact?\ny.Yes      n.No\n>>>";
                            while (true) {
                                getline(cin, command);
                                if (command == "y" || command == "Y") {
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
                                        cout << "\nContact deleted successfully!";
                                        cout << "\nPress any key to continue...";
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
                                        cout << "\n\nPress any key to continue...";
                                        while (true) {
                                            if (_kbhit()) {
                                                break;
                                            }
                                            Sleep(10);
                                        }
                                        break;
                                    }
                                }
                                else if (command == "n" || command == "N") {
                                    cout << "Delete operation canceled!";
                                    cout << "\nPress any key to continue...";
                                    while (true) {
                                        if (_kbhit()) {
                                            break;
                                        }
                                        Sleep(10);
                                    }
                                    break;
                                }
                                else {
                                    cout << "Invalid input, please try again!";
                                }
                            }
                            goto Da_Capo;
                        }
                    }
                    cout << "Contact not found";
                    cout << "\nPress any key to continue...";
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
        else if (command == "4") {
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