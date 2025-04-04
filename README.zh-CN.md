## 其他语言 Other languages
英语：[README](README.md)  
English:[README](README.md)
## 概述
你好啊！我是来自中国的一个学习C++不久的学生。这是一个非常简单的项目，我在练习使用xlnt库时写的  
它使用了xlnt库的表格读写，实现了制作并记录通讯录的功能  
## 功能
它在开始运行时会尝试打开 **"Memo.xlsx"** 表格，如果打开失败的话就会新建这个表格，并保存它
```cpp
try {
	xlnt::workbook wb;
	wb.load("Memo.xlsx");
}
catch (const std::exception& e) {
	xlnt::workbook wb;
	xlnt::worksheet ws = wb.sheet_by_title("Sheet1");
        //对于表格信息的一些初始化内容，在第一行写入“姓名、生日、住址、联系方式、备注”等内容
	wb.save("Memo.xlsx");
}
xlnt::workbook wb;
wb.load("Memo.xlsx");
xlnt::worksheet ws = wb.sheet_by_title("Sheet1");
```
运行程序后会有一个交互界面，可以输入 **"list"** 列出所有联系人，或者进入“写入联系人”的界面添加新的联系人，或者通过“查询联系人”、“删除联系人”界面输入名字查询、删除联系人  
每执行一次操作程序都会保存表格，以保证您的联系人信息不会丢失
## 运行
您可以将项目下载并使用 **Visual Studio** 打开 **"Memo_from_Friends.sln"** 直接编译运行，运行成功后可能会在 **"Memo_from_Friends.cpp"** 存在的目录下创建一个表格文档（如果您将原本的文档删除了的话），**请务必保护好它，它是您写入的所有联系人信息的唯一储存**
## 其他
### 读写中文
我作为一个中国人，首先当然要我的程序能够处理中文，所以在程序开头需要声明使用UTF-8编码(需要"Windows.h"这个头文件)，不然xlnt库无法处理中文，其实如果没有使用UTF-8编码的话它应该是只能读写ASCII码的
```cpp
#include <Windows.h>
int main(){
    SetConsoleOutputCP(65001);
    SetConsoleCP(65001);
}
```
然后需要在解决方案管理器右键源文件，点击 **属性→C/C++→命令行→其他选项** ，添加" **/utf-8** "
```
/utf-8
```
## 关于我
你好！我是一个正在学习C++的学生，有时可能会在GitHub分享一些在学习过程中写的代码，请多指教！  
### 我的名字
我的GitHub名字是"liuyingjiang-QwQ"，这是因为我非常喜欢中国米哈游（miHoYo）公司的一款游戏《崩坏：星穹铁道》中的一个叫“流萤”的角色，她的名字用中国的拼音拼过来就是"liuying"
### 联系我
我的邮箱就是：liuyingjiang_QwQ@outlook.com  
同时，我是中国的一个视频平台bilibili的一个小UP主，要不要来我主页看看呢？[bilibili主页](https://space.bilibili.com/3546591566760474)，我会在这里分享一些关于音乐、科技的内容
