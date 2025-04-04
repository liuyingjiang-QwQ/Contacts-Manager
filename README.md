
## Other languages 其他语言  
### README languages 自述文件语言
Chinese:[README.zh-CN](README.zh-CN.md)  
  
中文：[README.zh-CN](README.zh-CN.md)
### Projects languages 项目语言
English:[English Version](https://github.com/liuyingjiang-QwQ/Contacts-Manager/tree/English) (`English` branch)  
Chinese:[Chinese Version](https://github.com/liuyingjiang-QwQ/Contacts-Manager/tree/Chinese) (`Chinese` branch)  

英语：[英语版本](https://github.com/liuyingjiang-QwQ/Contacts-Manager/tree/English) (`English` 分支内)   
中文：[中文版本](https://github.com/liuyingjiang-QwQ/Contacts-Manager/tree/Chinese) (`Chinese` 分支内)  
## Overview  
Hello! I'm a student from China who is currently learning C++. This is a simple project I created while practicing with the xlnt library.  
It implements an address book system with Excel file read/write capabilities using xlnt.

## Features  
- The program attempts to open "Memo.xlsx" on startup. If the file doesn't exist, it creates a new one with initialized columns ("Name", "Birthday", "Address", "Contact", "Notes").
```cpp
try {
    xlnt::workbook wb;
    wb.load("Memo.xlsx");
}
catch (const std::exception& e) {
    xlnt::workbook wb;
    xlnt::worksheet ws = wb.sheet_by_title("Sheet1");
    // Initialize column headers...
    wb.save("Memo.xlsx");
}
```

- Interactive console interface with these commands:
  - `list`: Show all contacts
  - `Add Contact`: Add new contact
  - `Search Contact`: Find contact by name
  - `Delete Contact`: Remove contact
- Auto-saves after every operation to prevent data loss.

## Getting Started  
1. Select your preferred language (en-US or zh-CN)
2. Clone the repository
3. Open "Contacts-Manager.en-US.sln" in Visual Studio (If you need the Chinese version, please open "Contacts-Manager.zh-CN.sln”)
4. Build and run the project
5. The program will create "Memo.xlsx" in the same directory if it doesn't exist

> ⚠️ Important: This Excel file is the only storage for your contact data. Keep it safe!

## Technical Notes  
### UTF-8 Support  
To handle Chinese characters (and other non-ASCII text), the program requires:
1. UTF-8 console configuration:
```cpp
#include <Windows.h>
int main() {
    SetConsoleOutputCP(65001); // UTF-8
    SetConsoleCP(65001);
}
```
2. Compiler flag: Add `/utf-8` in:
   *solution Explorer → Source Files → \<your project> → Properties → C/C++ → Command Line → Additional Options*

## About Me  
I'm a C++ learner who occasionally shares study projects on GitHub. 

My GitHub username is "liuyingjiang-QwQ" because I love the character "Firefly" from *Honkai: Star Rail*, a game developed by the Chinese company HoYoverse. The character’s Chinese name "流萤" is phonetically spelled as "liuying" in Pinyin（A unique Chinese way of spelling）. 

### Contact  
- GitHub: [liuyingjiang-QwQ](https://github.com/liuyingjiang-QwQ)  
- Email: liuyingjiang_QwQ@outlook.com  
- Bilibili: [My Channel](https://space.bilibili.com/3546591566760474) (Chinese tech/music content)  
