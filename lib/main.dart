import 'package:flutter/material.dart';
import 'package:excel/excel.dart';
import 'package:open_file/open_file.dart';
import 'package:path_provider/path_provider.dart';
import 'custom_icon_button.dart';
import 'dart:io';
import 'package:intl/intl.dart';
import 'package:file_picker/file_picker.dart';
import 'package:path/path.dart' as path;

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({Key? key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Flutter Demo',
      theme: ThemeData(
        colorScheme: ColorScheme.fromSwatch().copyWith(secondary: Colors.deepPurple),
      ),
      home: const MyHomePage(title: 'Flutter Demo Home Page'),
    );
  }
}

class MyHomePage extends StatefulWidget {
  const MyHomePage({Key? key, required this.title}) : super(key: key);

  final String title;

  @override
  State<MyHomePage> createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  String? _filePath;

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      body: Center(
        child: Column(
          mainAxisAlignment: MainAxisAlignment.center,
          crossAxisAlignment: CrossAxisAlignment.stretch,
          children: [
            CustomIconButton(
              text: 'Export',
              onPressed: _onExportPressed,
              iconData: Icons.download,
            ),
            SizedBox(height: 40),
            CustomIconButton(
              text: 'Select Excel File',
              onPressed: _selectExcelFile,
              iconData: Icons.upload,
            ),
            SizedBox(height: 20),
            if (_filePath != null)
            Padding(padding: const EdgeInsets.all(16.0), child:
          Text("Seçilen dosya: ${path.basename(_filePath!)}",
        style: TextStyle(fontSize: 16, fontWeight: FontWeight.bold, color: Colors.deepPurple)),),

          ],
        ),
      ),
    );
  }

  void _onExportPressed() {
    createAndSaveExcel().then((_) {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Excel dosyası başarıyla kaydedildi!')),
      );
    }).catchError((error) {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Dosya kaydedilirken hata oluştu: $error')),
      );
    });
  }

  Future<void> createAndSaveExcel() async {
    try {
      var excel = Excel.createExcel();
      var sheet = excel['Puantaj'];

      List<CellValue?> dataList = [
        TextCellValue('No'),
        TextCellValue('Ad/Soyad'),
        TextCellValue('TC'),
        TextCellValue('Görevi'),
        TextCellValue('1'),
        TextCellValue('2'),
        TextCellValue('3'),
        TextCellValue('4'),
        TextCellValue('5'),
        TextCellValue('6'),
        TextCellValue('7'),
        TextCellValue('8'),
        TextCellValue('9'),
        TextCellValue('10'),
        TextCellValue('11'),
        TextCellValue('12'),
        TextCellValue('13'),
        TextCellValue('14'),
        TextCellValue('15'),
        TextCellValue('16'),
        TextCellValue('17'),
        TextCellValue('18'),
        TextCellValue('19'),
        TextCellValue('20'),
        TextCellValue('21'),
        TextCellValue('22'),
        TextCellValue('23'),
        TextCellValue('24'),
        TextCellValue('25'),
        TextCellValue('26'),
        TextCellValue('27'),
        TextCellValue('28'),
        TextCellValue('29'),
        TextCellValue('30'),
        TextCellValue('31')
      ];
      sheet.appendRow(dataList);

      var bytes = excel.save();

      final directory = await getExternalStorageDirectory(); // Harici depolama alanını alın
      if (directory == null) {
        throw 'Harici depolama alanı bulunamadı';
      }

      DateTime now = DateTime.now();
      String formattedDate = DateFormat('yyyy-MM-dd').format(now);
      String fileName = 'Puantaj-$formattedDate.xlsx';
      String filePath = '/storage/emulated/0/Download/$fileName';

      File file = File(filePath);

      await file.writeAsBytes(bytes!);
      print('Excel dosyası kaydedildi: $filePath');
    } catch (e) {
      print('Dosya kaydetme sırasında bir hata oluştu: $e');
    }
  }

  void readExcelFile(String filePath) async {
    var file = File(filePath);
    if (file.existsSync()) {
      var bytes = file.readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);

      for (var table in excel.tables.keys) {
        print(table); //sheet Name
        print(excel.tables[table]!.maxColumns);
        print(excel.tables[table]!.maxRows);
        for (var row in excel.tables[table]!.rows) {
          for (var cell in row) {
            print('cell ${cell?.rowIndex}/${cell?.columnIndex}');
            final value = cell?.value;
            print('Value: $value');
          }
        }
      }
    } else {
      print("Dosya bulunamadı!");
    }
  }

  Future<List<File>> listExcelFiles(String directoryPath) async {
    final directory = Directory(directoryPath);
    List<File> excelFiles = [];

    await for (var entity in directory.list()) {
      if (entity is File && entity.path.endsWith('.xlsx')) {
        excelFiles.add(entity);
      }
    }

    return excelFiles;
  }

  void _showExcelFilePicker(List<File> files) {
    showModalBottomSheet(
      context: context,
      builder: (context) {
        return ListView.builder(
          itemCount: files.length,
          itemBuilder: (context, index) {
            return ListTile(
              title: Text(path.basename(files[index].path)),
              onTap: () {
                Navigator.pop(context);
                setState(() {
                  _filePath = files[index].path;
                });
                readExcelFile(_filePath!);
              },
            );
          },
        );
      },
    );
  }

  void _selectExcelFile() async {
    const directoryPath = '/storage/emulated/0/Download/';
    List<File> files = await listExcelFiles(directoryPath);
    if (files.isNotEmpty) {
      OpenFile.open(files[0].path);
      return;
      setState(() {
        _showExcelFilePicker(files);
      });
    } else {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text('Dosya bulunamadı')),
      );
    }
  }
}