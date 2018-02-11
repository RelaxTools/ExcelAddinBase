# ExcelAddinBase
Excel 2010/2013/2016 Ribbon対応版
Excel Addin Platform
-------------------------------------------
Excel アドイン作成ベースです。ご自由にどうぞ。

◇使用方法
customUi.xml はレイアウト専用。
![customUI](https://github.com/RelaxTools/ExcelAddinBase/blob/master/media/customUi.png)

メニューの文字列やアイコンを変更する場合、ThisWorkbookのプロパティIsAddin=Falseにする。
![IsAddin](https://github.com/RelaxTools/ExcelAddinBase/blob/master/media/IsAddin.png)

中のシートが表示されるので変更する。IsAddin=Trueにして保存する。
![Sheet](https://github.com/RelaxTools/ExcelAddinBase/blob/master/media/Sheet.png)

ボタンとマクロを紐付けるにはIDと同じマクロ名を作成する。
![Module](https://github.com/RelaxTools/ExcelAddinBase/blob/master/media/Module.png)

