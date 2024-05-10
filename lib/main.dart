import 'package:flutter/material.dart';
import 'package:excel/excel.dart';
import 'package:path_provider/path_provider.dart';
import 'custom_icon_button.dart';
import 'dart:io';

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
      var sheet = excel['Sheet1'];

      List<CellValue?> dataList = [
        TextCellValue('Google'),
        TextCellValue('loves'),
        TextCellValue('Flutter'),
        TextCellValue('and'),
        TextCellValue('Flutter'),
        TextCellValue('loves'),
        TextCellValue('Excel')
      ];
      sheet.appendRow(dataList);

      var bytes = excel.save();

      final directory = await getExternalStorageDirectory(); // Harici depolama alanını alın
      if (directory == null) {
        throw 'Harici depolama alanı bulunamadı';
      }

      String filePath = '/storage/emulated/0/Download/Deneme.xlsx'; // Belirtilen klasöre kaydedin
      File file = File(filePath);

      await file.writeAsBytes(bytes!);
      print('Excel dosyası kaydedildi: $filePath');
    } catch (e) {
      print('Dosya kaydetme sırasında bir hata oluştu: $e');
    }
  }
}