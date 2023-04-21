import 'package:flutter/material.dart';
import 'dart:async';

import 'package:dartxlsxwriter/dartxlsxwriter.dart' as dartxlsxwriter;

void main() {
  runApp(const MyApp());
}

class MyApp extends StatefulWidget {
  const MyApp({super.key});

  @override
  State<MyApp> createState() => _MyAppState();
}

class _MyAppState extends State<MyApp> {
  // late int sumResult;
  // late Future<int> sumAsyncResult;

  @override
  void initState() {
    super.initState();
    var thing = dartxlsxwriter.Workbook('filename.xlsx');
    thing.newWorksheet('sheetname');
    thing.worksheets['sheetname']!.writeValue(1, 1, 1);
    thing.worksheets['sheetname']!.writeValue(0, 0,
        'this file was created by libxlsxwriter running via flutter / dart');
    thing.worksheets['sheetname']!.writeValue(2, 0,
        'this file was created by libxlsxwriter running via flutter / dart');
    thing.close();
    // sumResult = dartxlsxwriter.
    // sumAsyncResult = dartxlsxwriter.sumAsync(3, 4);
  }

  @override
  Widget build(BuildContext context) {
    const textStyle = TextStyle(fontSize: 25);
    const spacerSmall = SizedBox(height: 10);
    return MaterialApp(
      home: Scaffold(
        appBar: AppBar(
          title: const Text('Native Packages'),
        ),
        body: SingleChildScrollView(
          child: Container(
            padding: const EdgeInsets.all(10),
            child: Column(
              children: const [
                Text(
                  'This calls a native function through FFI that is shipped as source in the package. '
                  'The native code is built as part of the Flutter Runner build.',
                  style: textStyle,
                  textAlign: TextAlign.center,
                ),
                spacerSmall,
                Text(
                  'file saved',
                  style: textStyle,
                  textAlign: TextAlign.center,
                ),
                spacerSmall,
              ],
            ),
          ),
        ),
      ),
    );
  }
}
