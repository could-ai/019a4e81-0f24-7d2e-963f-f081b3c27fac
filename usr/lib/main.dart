import 'package:flutter/material.dart';
import 'package:flutter/services.dart' show ByteData, rootBundle;
import 'package:excel/excel.dart';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Excel Reader Demo',
      debugShowCheckedModeBanner: false,
      theme: ThemeData(
        primarySwatch: Colors.blue,
        useMaterial3: true,
      ),
      home: const ExcelReaderPage(),
    );
  }
}

class ExcelReaderPage extends StatefulWidget {
  const ExcelReaderPage({super.key});

  @override
  State<ExcelReaderPage> createState() => _ExcelReaderPageState();
}

class _ExcelReaderPageState extends State<ExcelReaderPage> {
  List<List<Data?>>? _excelData;
  String? _error;

  @override
  void initState() {
    super.initState();
    _loadExcelFile();
  }

  Future<void> _loadExcelFile() async {
    try {
      // Load the Excel file from assets
      ByteData data = await rootBundle.load("assets/data.xlsx");
      var bytes = data.buffer.asUint8List(data.offsetInBytes, data.lengthInBytes);
      var excel = Excel.decodeBytes(bytes);

      if (excel.tables.keys.isNotEmpty) {
        String sheetName = excel.tables.keys.first;
        var sheet = excel.tables[sheetName];
        if (sheet != null) {
          setState(() {
            _excelData = sheet.rows;
          });
        } else {
           setState(() {
            _error = "Sheet not found.";
          });
        }
      } else {
         setState(() {
            _error = "No sheets found in the Excel file.";
          });
      }
    } catch (e) {
      setState(() {
        _error = "Error loading or parsing Excel file: $e\n\nPlease ensure you have a file named 'data.xlsx' inside an 'assets' folder in your project root.";
      });
      print(e.toString());
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text("Excel Data Reader"),
      ),
      body: Center(
        child: _buildContent(),
      ),
    );
  }

  Widget _buildContent() {
    if (_error != null) {
      return Padding(
        padding: const EdgeInsets.all(16.0),
        child: Text(
          _error!,
          style: const TextStyle(color: Colors.red, fontSize: 16),
          textAlign: TextAlign.center,
        ),
      );
    }

    if (_excelData == null) {
      return const CircularProgressIndicator();
    }

    if (_excelData!.isEmpty) {
      return const Text("The Excel sheet is empty.");
    }

    // Assuming the first row is the header
    List<Data?> headerRow = _excelData![0];
    List<List<Data?>> dataRows = _excelData!.sublist(1);

    return SingleChildScrollView(
      scrollDirection: Axis.vertical,
      child: SingleChildScrollView(
        scrollDirection: Axis.horizontal,
        child: DataTable(
          columns: headerRow.map((cell) {
            return DataColumn(
              label: Text(
                cell?.value?.toString() ?? '',
                style: const TextStyle(fontWeight: FontWeight.bold),
              ),
            );
          }).toList(),
          rows: dataRows.map((row) {
            return DataRow(
              cells: row.map((cell) {
                return DataCell(Text(cell?.value?.toString() ?? ''));
              }).toList(),
            );
          }).toList(),
        ),
      ),
    );
  }
}
