# ExcelAddinBase
Excel 2010/2013/2016 Ribbon対応版
Excel Addin Platform
-------------------------------------------
Excel アドイン作成ベースです。ご自由にどうぞ。

◇使用方法
customUi.xml はレイアウト専用。
![](https://github.com/RelaxTools/ExcelAddinBase/media/customUi.png)

メニューの文字列やアイコンを変更する場合、ThisWorkbookのプロパティIsAddin=Falseにする。
![](https://github.com/RelaxTools/ExcelAddinBase/media/IsAddin.png)

中のシートが表示されるので変更する。IsAddin=Trueにして保存する。
![](https://github.com/RelaxTools/ExcelAddinBase/media/Sheet.png)

ボタンとマクロを紐付けるにはIDと同じマクロ名を作成する。
![](https://github.com/RelaxTools/ExcelAddinBase/media/Module.png)

