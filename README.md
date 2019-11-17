# VBA

Excel裏面沒有註冊 Microsoft Date and Time Picker Control 6.0（SP6）控件。

## 下載控件手動註冊後解決問題。

OS作業系統：Windows 10 x64
Office版本：2016 x32 ， 2019 x32 (只能32位)

Step 01： 下載控件：https://github.com/ableshih/VBA/blob/master/MSCOMCT2.zip 

Step 02： 解壓。並將其中的“mscomct2.ocx”文件複製到“\Windows\SysWOW64”文件夾中。

Step 03： 以系統管理員身分執行 “命令提示字元”，輸入 regsvr32 "\Windows\SysWOW64\mscomct2.ocx”

Step 04： 系統提示註冊成功。

來源
http://www.logicwurks.com/CodeExamplePages/EDatePickerControl.html
http://activex.microsoft.com/controls/vb6/mscomct2.cab
