import 'dart:io';

import 'package:numerus/numerus.dart';
import 'package:flutter/material.dart';
import 'package:csv/csv.dart';
import 'dart:async' show Future;
import 'package:flutter/services.dart' show rootBundle;
import 'package:syncfusion_flutter_xlsio/xlsio.dart' as xlsio1;
import 'package:path_provider/path_provider.dart';
import 'package:universal_html/html.dart' show AnchorElement;
import 'package:flutter/foundation.dart' show kIsWeb;
import 'dart:convert';

void main() => runApp(MyApp());

class MyApp extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Генератор excel',
      home: HomePage(),
    );
  }
}

class HomePage extends StatefulWidget {
  @override
  _HomePageState createState() => _HomePageState();
}

class _HomePageState extends State<HomePage> {
  List<dynamic> jsonData = [];
  List<dynamic> filteredData = [];
  List<bool> checkedList = [];
  List<int> filteredIndices = [];

  @override
  void initState() {
    super.initState();
    loadAsset().then((value) {
      setState(() {
        jsonData = json.decode(value);
        filteredData = List.from(jsonData);
        checkedList = List.filled(jsonData.length, false);
        filteredIndices =
            List<int>.generate(jsonData.length, (i) => i); // Add this line
      });
    });
  }

  // Future<String> loadAsset() async {
  //   return await rootBundle.loadString('lib/data.csv');
  // }

  int getSelectedCount() {
    return checkedList.where((isChecked) => isChecked).length;
  }

  Future<String> loadAsset() async {
    return await rootBundle.loadString('lib/TableData.json');
  }

  // void filterData(String query) {
  //   setState(() {
  //     if (query.isEmpty) {
  //       filteredData = List.from(jsonData);
  //     } else {
  //       filteredData = jsonData
  //           .where((data) =>
  //               data['Name'].toLowerCase().contains(query.toLowerCase()))
  //           .toList();
  //     }
  //   });
  // }

  void filterData(String query) {
    setState(() {
      if (query.isEmpty) {
        filteredData = List.from(jsonData);
        filteredIndices = List<int>.generate(jsonData.length, (i) => i);
      } else {
        filteredData = [];
        filteredIndices = [];
        for (int i = 0; i < jsonData.length; i++) {
          if (jsonData[i]['Name'].toLowerCase().contains(query.toLowerCase())) {
            filteredData.add(jsonData[i]);
            filteredIndices.add(i);
          }
        }
      }
    });
  }

  void toggleCheckbox(int index) {
    int originalIndex = filteredIndices[index];
    setState(() {
      checkedList[originalIndex] = !checkedList[originalIndex];
    });
  }

  void showCheckedOptions() {
    List<String> selectedOptions = [];
    for (int i = 0; i < filteredData.length; i++) {
      if (checkedList[i]) {
        selectedOptions.add(filteredData[i][0].toString());
      }
    }
    showDialog(
      context: context,
      builder: (BuildContext context) {
        return AlertDialog(
          title: Text('Selected options'),
          content: Text(selectedOptions.join('\n')),
          actions: [
            TextButton(
              onPressed: () => Navigator.pop(context),
              child: Text('OK'),
            ),
          ],
        );
      },
    );
  }

  Future<void> _createExcel() async {
//Create an Excel document.

    //Creating a workbook.
    final xlsio1.Workbook workbook = xlsio1.Workbook();
    //Accessing via index
    final xlsio1.Worksheet sheet = workbook.worksheets[0];
    sheet.getRangeByName('B14:G100').cellStyle.wrapText = true;
    sheet.getRangeByName('A1:G100').cellStyle.fontName = "Times New Roman";
    sheet.getRangeByName('A14:G100').cellStyle.fontSize = 14;
    sheet.getRangeByName('B19:E100').cellStyle.hAlign = xlsio1.HAlignType.left;
    sheet.getRangeByName('A17:G17').cellStyle.hAlign = xlsio1.HAlignType.center;
    sheet.getRangeByName('A17:G17').cellStyle.vAlign = xlsio1.VAlignType.center;

    // sheet.getRangeByName('A14:G100').cellStyle.vAlign = xlsio1.VAlignType.top;
    // sheet.getRangeByName('A14:G17').cellStyle.hAlign = xlsio1.HAlignType.center;
    // sheet.getRangeByName('G18:E100').cellStyle.hAlign = xlsio1.HAlignType.left;

    sheet.getRangeByName('A1').columnWidth = 10;
    sheet.getRangeByName('B1').columnWidth = 35;
    sheet.getRangeByName('C1').columnWidth = 35;
    sheet.getRangeByName('D1').columnWidth = 35;
    sheet.getRangeByName('E1').columnWidth = 25;
    sheet.getRangeByName('F1').columnWidth = 19;
    sheet.getRangeByName('G1').columnWidth = 19;

    sheet.getRangeByName('H1:K1000').cells.remove;

    sheet.getRangeByName('B14:D15').cellStyle.borders.all.lineStyle =
        xlsio1.LineStyle.thin;

    sheet
        .getRangeByIndex(17, 1, 18 + getSelectedCount(), 7)
        .cellStyle
        .borders
        .all
        .lineStyle = xlsio1.LineStyle.thin;

    //     sheet.getRangeByName('A17:G100').cellStyle.borders.all.lineStyle =
    // xlsio1.LineStyle.thin;

    sheet.getRangeByName('F1:G1').merge();
    sheet.getRangeByName('F1:G1').setText("Приложение № 2");
    sheet.getRangeByName('F1:G1').cellStyle.fontSize = 14;

    sheet.getRangeByName('F2:G2').merge();
    sheet.getRangeByName('F2:G2').setText("к приказу Администрации");
    sheet.getRangeByName('F2:G2').cellStyle.fontSize = 14;
    sheet.getRangeByName('F1:G1').cellStyle.hAlign = xlsio1.HAlignType.right;

    sheet.getRangeByName('F3:G3').merge();
    sheet.getRangeByName('F3:G3').setText("Главы Республики Коми");
    sheet.getRangeByName('F3:G3').cellStyle.fontSize = 14;

    sheet.getRangeByName('F4:G4').merge();
    sheet.getRangeByName('F4:G4').setText("от «___» ________2020 г. №__");
    sheet.getRangeByName('F4:G4').cellStyle.fontSize = 14;

    sheet.getRangeByName('F5:G5').merge();
    sheet.getRangeByName('F5:G5').setText("(форма)	");
    sheet.getRangeByName('F5:G5').cellStyle.fontSize = 14;

    sheet.getRangeByName('F1:G5').cellStyle.hAlign = xlsio1.HAlignType.right;

    sheet.getRangeByName('A6:G6').merge();
    sheet.getRangeByName('A6:G6').setText("ИНДИВИДУАЛЬНЫЙ ПЛАН");
    sheet.getRangeByName('A6:G6').cellStyle.hAlign = xlsio1.HAlignType.center;
    sheet.getRangeByName('A6:G6').cellStyle.bold = true;
    sheet.getRangeByName('A6:G6').cellStyle.fontSize = 14;

    sheet.getRangeByName('A7:G7').merge();
    sheet.getRangeByName('A7:G7').setText(
        "профессионального развития государственного гражданского служащего");
    sheet.getRangeByName('A7:G7').cellStyle.bold = true;
    sheet.getRangeByName('A7:G7').cellStyle.hAlign = xlsio1.HAlignType.center;
    sheet.getRangeByName('A7:G7').cellStyle.fontSize = 14;

    sheet.getRangeByName('A8:G8').merge();
    sheet.getRangeByName('A8:G8').setText(_nameController.text);
    sheet.getRangeByName('A8:G8').cellStyle.hAlign = xlsio1.HAlignType.center;
    sheet.getRangeByName('A8:G8').cellStyle.fontSize = 14;
    sheet.getRangeByName('B8:G8').cellStyle.borders.bottom.lineStyle =
        xlsio1.LineStyle.thin;

    sheet.getRangeByName('A9:G9').merge();
    sheet.getRangeByName('A9:G9').setText("(ФИО)");
    sheet.getRangeByName('A9:G9').cellStyle.hAlign = xlsio1.HAlignType.center;

    sheet.getRangeByName('A10:G10').merge();
    sheet.getRangeByName('A10:G10').setText(_dolzhController.text);
    sheet.getRangeByName('A10:G10').cellStyle.hAlign = xlsio1.HAlignType.center;
    sheet.getRangeByName('A10:G10').cellStyle.fontSize = 14;
    sheet.getRangeByName('B10:G10').cellStyle.borders.bottom.lineStyle =
        xlsio1.LineStyle.thin;

    sheet.getRangeByName('A11:G11').merge();
    sheet.getRangeByName('A11:G11').setText(
        "(должность с указанием наименования структурного подразделения, государственного органа)");
    sheet.getRangeByName('A11:G11').cellStyle.hAlign = xlsio1.HAlignType.center;

    sheet.getRangeByName('C12').setText("на");
    sheet.getRangeByName('C12').cellStyle.hAlign = xlsio1.HAlignType.right;

    sheet.getRangeByName('D12').setText(_yearController.text);
    sheet.getRangeByName('D12').cellStyle.hAlign = xlsio1.HAlignType.center;
    sheet.getRangeByName('D12').cellStyle.borders.bottom.lineStyle =
        xlsio1.LineStyle.thin;

    sheet.getRangeByName('E12').setText("годы");
    sheet.getRangeByName('E12').cellStyle.hAlign = xlsio1.HAlignType.left;

    sheet
        .getRangeByName('B14')
        .setText("Область профессиональной служебной деятельности");

    sheet
        .getRangeByName('C14')
        .setText("Вид профессиональной служебной деятельности");
    sheet.getRangeByName('D14').setText("Функциональные обязанности");

    sheet.getRangeByName('B14:D15').cellStyle.vAlign = xlsio1.VAlignType.top;
    sheet.getRangeByName('B14:D15').cellStyle.hAlign = xlsio1.HAlignType.center;

    sheet.getRangeByName('A17').setText("№ пп");
    sheet
        .getRangeByName('B17')
        .setText("Цель профессионального развития, ожидаемая результативность");
    sheet
        .getRangeByName('C17')
        .setText("Направление профессионального развития ");
    sheet
        .getRangeByName('D17')
        .setText("Тема мероприятия по профессиональному развитию");
    sheet
        .getRangeByName('E17')
        .setText("Вид/ форма  профессионального развития");
    sheet.getRangeByName('F17').setText("Продолжительность (в часах)");
    sheet.getRangeByName('G17').setText("Сроки (год)");

    sheet.getRangeByName('A18:G18').cellStyle.hAlign = xlsio1.HAlignType.center;
    sheet.getRangeByName('A18').setText("1");
    sheet.getRangeByName('B18').setText("2");
    sheet.getRangeByName('C18').setText("3");
    sheet.getRangeByName('D18').setText("4");
    sheet.getRangeByName('E18').setText("5");
    sheet.getRangeByName('F18').setText("6");
    sheet.getRangeByName('G18').setText("7");

    sheet.getRangeByIndex(19 + getSelectedCount(), 2).setText("Исполнитель:");
    sheet.getRangeByIndex(19 + getSelectedCount(), 2).cellStyle.fontSize = 14;

    sheet
        .getRangeByIndex(20 + getSelectedCount(), 2)
        .cellStyle
        .borders
        .bottom
        .lineStyle = xlsio1.LineStyle.thin;

    sheet.getRangeByIndex(21 + getSelectedCount(), 2).setText("(ФИО)");
    sheet.getRangeByIndex(21 + getSelectedCount(), 2).cellStyle.hAlign =
        xlsio1.HAlignType.center;
    sheet.getRangeByIndex(21 + getSelectedCount(), 2).cellStyle.vAlign =
        xlsio1.VAlignType.top;

    sheet.getRangeByIndex(21 + getSelectedCount(), 2).cellStyle.fontSize = 9;
    sheet.getRangeByIndex(22 + getSelectedCount(), 2).setText("Согласовано:");

    sheet
        .getRangeByIndex(23 + getSelectedCount(), 2)
        .cellStyle
        .borders
        .bottom
        .lineStyle = xlsio1.LineStyle.thin;

    sheet.getRangeByIndex(24 + getSelectedCount(), 2).setText(
        "(наименование должности непосредственного руководителя гражданского служащего)");
    sheet.getRangeByIndex(24 + getSelectedCount(), 2).cellStyle.fontSize = 9;
    sheet.getRangeByIndex(24 + getSelectedCount(), 2).cellStyle.hAlign =
        xlsio1.HAlignType.center;

    sheet
        .getRangeByIndex(26 + getSelectedCount(), 2)
        .cellStyle
        .borders
        .bottom
        .lineStyle = xlsio1.LineStyle.thin;

    sheet.getRangeByIndex(27 + getSelectedCount(), 2).setText(
        "(наименование должности руководителя кадровой службы (либо должностного лица, ответственного за ведение кадровой работы) органа государственной власти)");
    sheet.getRangeByIndex(27 + getSelectedCount(), 2).cellStyle.fontSize = 9;
    sheet.getRangeByIndex(27 + getSelectedCount(), 2).cellStyle.hAlign =
        xlsio1.HAlignType.center;
    sheet.getRangeByIndex(27 + getSelectedCount(), 2).cellStyle.vAlign =
        xlsio1.VAlignType.top;

    sheet
        .getRangeByIndex(20 + getSelectedCount(), 4)
        .cellStyle
        .borders
        .bottom
        .lineStyle = xlsio1.LineStyle.thin;
    sheet.getRangeByIndex(21 + getSelectedCount(), 4).setText("(подпись)");
    sheet.getRangeByIndex(21 + getSelectedCount(), 4).cellStyle.fontSize = 9;
    sheet.getRangeByIndex(21 + getSelectedCount(), 4).cellStyle.hAlign =
        xlsio1.HAlignType.center;
    sheet.getRangeByIndex(21 + getSelectedCount(), 4).cellStyle.vAlign =
        xlsio1.VAlignType.top;

    sheet
        .getRangeByIndex(23 + getSelectedCount(), 4)
        .cellStyle
        .borders
        .bottom
        .lineStyle = xlsio1.LineStyle.thin;

    sheet.getRangeByIndex(24 + getSelectedCount(), 4).setText("(подпись)");
    sheet.getRangeByIndex(24 + getSelectedCount(), 4).cellStyle.fontSize = 9;
    sheet.getRangeByIndex(24 + getSelectedCount(), 4).cellStyle.hAlign =
        xlsio1.HAlignType.center;
    sheet.getRangeByIndex(24 + getSelectedCount(), 4).cellStyle.vAlign =
        xlsio1.VAlignType.top;

    sheet
        .getRangeByIndex(26 + getSelectedCount(), 4)
        .cellStyle
        .borders
        .bottom
        .lineStyle = xlsio1.LineStyle.thin;
    sheet.getRangeByIndex(27 + getSelectedCount(), 4).setText("(подпись)");
    sheet.getRangeByIndex(27 + getSelectedCount(), 4).cellStyle.fontSize = 9;
    sheet.getRangeByIndex(27 + getSelectedCount(), 4).cellStyle.hAlign =
        xlsio1.HAlignType.center;
    sheet.getRangeByIndex(27 + getSelectedCount(), 4).cellStyle.vAlign =
        xlsio1.VAlignType.top;

    sheet
        .getRangeByIndex(23 + getSelectedCount(), 6, 23 + getSelectedCount(), 7)
        .merge();
    sheet
        .getRangeByIndex(24 + getSelectedCount(), 6, 24 + getSelectedCount(), 7)
        .merge();
    sheet
        .getRangeByIndex(26 + getSelectedCount(), 6, 26 + getSelectedCount(), 7)
        .cellStyle
        .borders
        .bottom
        .lineStyle = xlsio1.LineStyle.thin;

    sheet
        .getRangeByIndex(26 + getSelectedCount(), 6, 26 + getSelectedCount(), 7)
        .merge();
    sheet
        .getRangeByIndex(27 + getSelectedCount(), 6, 27 + getSelectedCount(), 7)
        .merge();
    sheet
        .getRangeByIndex(27 + getSelectedCount(), 6, 27 + getSelectedCount(), 7)
        .setText("(ФИО)");
    sheet
        .getRangeByIndex(27 + getSelectedCount(), 6, 27 + getSelectedCount(), 7)
        .cellStyle
        .fontSize = 9;
    sheet
        .getRangeByIndex(27 + getSelectedCount(), 6, 27 + getSelectedCount(), 7)
        .cellStyle
        .hAlign = xlsio1.HAlignType.center;
    sheet
        .getRangeByIndex(27 + getSelectedCount(), 6, 27 + getSelectedCount(), 7)
        .cellStyle
        .vAlign = xlsio1.VAlignType.top;

    sheet
        .getRangeByIndex(23 + getSelectedCount(), 6, 23 + getSelectedCount(), 7)
        .cellStyle
        .borders
        .bottom
        .lineStyle = xlsio1.LineStyle.thin;

    sheet.getRangeByIndex(24 + getSelectedCount(), 6).setText("(ФИО)");
    sheet.getRangeByIndex(24 + getSelectedCount(), 6).cellStyle.fontSize = 9;
    sheet.getRangeByIndex(24 + getSelectedCount(), 6).cellStyle.hAlign =
        xlsio1.HAlignType.center;
    sheet.getRangeByIndex(24 + getSelectedCount(), 6).cellStyle.vAlign =
        xlsio1.VAlignType.top;

    int num1 = 18;

    for (int i = 0; i < filteredData.length; i++) {
      if (checkedList[i]) {
        num1++;
        String stringNum = num1.toString();
        sheet.getRangeByName('B' + stringNum).setText(filteredData[i]['Goal']);
        sheet.getRangeByName('B' + stringNum).cellStyle.vAlign =
            xlsio1.VAlignType.top;
        sheet
            .getRangeByName('C' + stringNum)
            .setText(filteredData[i]['Direction']);
        sheet.getRangeByName('D' + stringNum).setText(filteredData[i]['Name']);
        sheet.getRangeByName('E' + stringNum).setText(filteredData[i]['Type']);
        sheet
            .getRangeByName('F' + stringNum)
            .setText(filteredData[i]['Duration']);
      }
    }

    sheet
        .getRangeByName('C18:E' + (getSelectedCount() + 18).toString())
        .cellStyle
        .vAlign = xlsio1.VAlignType.top;

    sheet
        .getRangeByName('F18:G' + (getSelectedCount() + 18).toString())
        .cellStyle
        .vAlign = xlsio1.VAlignType.top;

    sheet
        .getRangeByName('F18:G' + (getSelectedCount() + 18).toString())
        .cellStyle
        .hAlign = xlsio1.HAlignType.center;

    int howManyMerged = 0;
    int howManyMergedCol3 = 0;

    String? nextTarget;
    String? currentTarget;

    String? nextTargetCol3;
    String? currentTargetCol3;

    int number = 1;
    int numStart = 18;
    for (int i = 0; i < getSelectedCount(); i++) {
      numStart++;
      // if (sheet.getRangeByIndex(numStart, 2).getText() ==
      //     sheet.getRangeByIndex(numStart + 1, 2).getText()) {
      //   sheet.getRangeByIndex(numStart, 2, numStart + 1, 2).merge();
      // }
      currentTarget = sheet.getRangeByIndex(numStart, 2).getText();
      nextTarget = sheet.getRangeByIndex(numStart + 1, 2).getText();

      nextTargetCol3 = sheet.getRangeByIndex(numStart, 3).getText();
      currentTargetCol3 = sheet.getRangeByIndex(numStart + 1, 3).getText();

      if (nextTarget == currentTarget) {
        howManyMerged++;
        if (nextTargetCol3 == currentTargetCol3) {
          howManyMergedCol3++;
        } else {
          if (howManyMergedCol3 > 0) {
            sheet
                .getRangeByIndex(numStart - howManyMergedCol3, 3, numStart, 3)
                .merge();
            howManyMergedCol3 = 0;
          }
        }
      } else {
        sheet.getRangeByIndex(numStart - howManyMerged, 2, numStart, 2).merge();
        sheet.getRangeByIndex(numStart - howManyMerged, 1, numStart, 1).merge();
        sheet
            .getRangeByIndex(numStart - howManyMerged, 1)
            .setText(number.toRomanNumeralString()! + ".");
        sheet.getRangeByIndex(numStart - howManyMerged, 1).cellStyle.vAlign =
            xlsio1.VAlignType.center;
        sheet.getRangeByIndex(numStart - howManyMerged, 1).cellStyle.hAlign =
            xlsio1.HAlignType.center;

        number++;
        howManyMerged = 0;
        if (howManyMergedCol3 > 0) {
          sheet
              .getRangeByIndex(numStart - howManyMergedCol3, 3, numStart, 3)
              .merge();
          howManyMergedCol3 = 0;
        }
      }
    }

    AnchorElement(
        href:
            'data:application/octet-stream;base64,${base64.encode(workbook.saveAsStream())}')
      ..setAttribute('download', 'sample.xlsx')
      ..click();
  }

  final _nameController = TextEditingController();
  final _yearController = TextEditingController();
  final _dolzhController = TextEditingController();

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      // floatingActionButton: FloatingActionButton(
      //   onPressed: () => _createExcel(),
      //   child: Icon(Icons.check),
      // ),
      body: Column(
        children: [
          SizedBox(
            height: 15,
          ),
          SizedBox(
            width: 450,
            child: TextField(
              controller: _nameController,
              decoration: InputDecoration(
                border: OutlineInputBorder(),
                labelText: 'Полное ФИО',
              ),
            ),
          ),
          SizedBox(height: 16),
          SizedBox(
            width: 450,
            child: TextField(
              controller: _dolzhController,
              decoration: InputDecoration(
                border: OutlineInputBorder(),
                labelText:
                    'Должность с указанием наименования структурного подразделения, государственного органа',
              ),
            ),
          ),
          SizedBox(height: 16),
          SizedBox(
            width: 450,
            child: TextField(
              controller: _yearController,
              decoration: InputDecoration(
                border: OutlineInputBorder(),
                labelText: 'Годы обучения',
              ),
            ),
          ),
          SizedBox(height: 16),
          SizedBox(
            width: 600,
            child: TextField(
              onChanged: (value) => filterData(value),
              decoration: InputDecoration(
                hintText: 'Поиск по названию программы',
                border: UnderlineInputBorder(),
                contentPadding: EdgeInsets.symmetric(horizontal: 16),
              ),
            ),
          ),
          SizedBox(height: 16),
          Expanded(
            child: ListView.builder(
              itemCount: filteredData.length,
              itemBuilder: (BuildContext context, int index) {
                int originalIndex = filteredIndices[index];
                return ListTile(
                  title: Text(filteredData[index]['Name']),
                  subtitle: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Text(filteredData[index]['Goal']),
                      Text(filteredData[index]['Direction']),
                      Text(filteredData[index]['Type']),
                      Text(filteredData[index]['Duration'] + ' ак. час.'),
                    ],
                  ),
                  trailing: Checkbox(
                    value: checkedList[originalIndex],
                    onChanged: (value) => toggleCheckbox(index),
                  ),
                );
              },
            ),
          ),
          SizedBox(height: 16),
          SizedBox(
              height: 50,
              width: 450,
              child: ElevatedButton(
                  onPressed: () {
                    _createExcel();
                  },
                  child: Text("Сгенерировать таблицу"))),
          SizedBox(height: 16)
        ],
      ),
    );
  }
}
