import 'dart:async';
import 'dart:convert';
import 'dart:ffi';
import 'package:ffi/ffi.dart';
import 'dart:io';
import 'dart:isolate';

import 'dartxlsxwriter_bindings_generated.dart';

/// A very short-lived native function.
///
/// For very short-lived functions, it is fine to call them on the main isolate.
/// They will block the Dart execution while running the native function, so
/// only do this for native functions which are guaranteed to be short-lived.
// int sum(int a, int b) => _bindings.sum(a, b);

/// A longer lived native function, which occupies the thread calling it.
///
/// Do not call these kind of native functions in the main isolate. They will
/// block Dart execution. This will cause dropped frames in Flutter applications.
/// Instead, call these native functions on a separate isolate.
///
/// Modify this to suit your own use case. Example use cases:
///
/// 1. Reuse a single isolate for various different kinds of requests.
/// 2. Use multiple helper isolates for parallel execution.
// Future<int> sumAsync(int a, int b) async {
//   final SendPort helperIsolateSendPort = await _helperIsolateSendPort;
//   final int requestId = _nextSumRequestId++;
//   final _SumRequest request = _SumRequest(requestId, a, b);
//   final Completer<int> completer = Completer<int>();
//   _sumRequests[requestId] = completer;
//   helperIsolateSendPort.send(request);
//   return completer.future;
// }

const String _libName = 'dartxlsxwriter';

/// The dynamic library in which the symbols for [DartxlsxwriterBindings] can be found.
final DynamicLibrary _dylib = () {
  if (Platform.isMacOS || Platform.isIOS) {
    return DynamicLibrary.open('$_libName.framework/$_libName');
  }
  if (Platform.isAndroid || Platform.isLinux) {
    return DynamicLibrary.open('lib$_libName.so');
  }
  if (Platform.isWindows) {
    return DynamicLibrary.open('$_libName.dll');
  }
  throw UnsupportedError('Unknown platform: ${Platform.operatingSystem}');
}();

/// The bindings to the native functions in [_dylib].
final DartxlsxwriterBindings _bindings = DartxlsxwriterBindings(_dylib);

class Worksheet {
  final Pointer<lxw_worksheet> ws;
  final String worksheetname;
  final Map<int, Map<int, dynamic>> cells = {};
  Worksheet(this.worksheetname, this.ws);

  void writeValue(int row, int column, dynamic value) {
    switch (value.runtimeType) {
      case int:
        print(_bindings.worksheet_write_number(
            ws, row, column, value.toDouble(), nullptr));
        break;
      case double:
        _bindings.worksheet_write_number(ws, row, column, value, nullptr);
        break;
      case String:
        final String convertString = value.toString();
        final stringChar = convertString.toNativeUtf8().cast<Char>();
        _bindings.worksheet_write_string(ws, row, column, stringChar, nullptr);
        break;
      default:
        try {
          final stringChar = value.toString().toNativeUtf8().cast<Char>();
          _bindings.worksheet_write_string(
              ws, row, column, stringChar, nullptr);
        } catch (e) {
          throw Exception('Unimplemented type');
        }
        break;
    }
    cells.putIfAbsent(row, () => {});
    cells[row]![column] = value;
  }
}

class Workbook {
  late Pointer<lxw_workbook> wb;
  final String filename;
  final Map<String, Worksheet> worksheets = {};
  Workbook(this.filename) {
    final filenamePtr = filename.toNativeUtf8().cast<Char>();
    wb = _bindings.workbook_new(filenamePtr);
  }

  void newWorksheet(String sheetname) {
    final sheetnamePtr = sheetname.toNativeUtf8().cast<Char>();
    worksheets[sheetname] = Worksheet(
        sheetname, _bindings.workbook_add_worksheet(wb, sheetnamePtr));
  }

  void close() {
    print(_bindings.workbook_close(wb));
    wb = nullptr;
  }
}
// /// A request to compute `sum`.
// ///
// /// Typically sent from one isolate to another.
// class _SumRequest {
//   final int id;
//   final int a;
//   final int b;

//   const _SumRequest(this.id, this.a, this.b);
// }

// /// A response with the result of `sum`.
// ///
// /// Typically sent from one isolate to another.
// class _SumResponse {
//   final int id;
//   final int result;

//   const _SumResponse(this.id, this.result);
// }

// /// Counter to identify [_SumRequest]s and [_SumResponse]s.
// int _nextSumRequestId = 0;

// /// Mapping from [_SumRequest] `id`s to the completers corresponding to the correct future of the pending request.
// final Map<int, Completer<int>> _sumRequests = <int, Completer<int>>{};

// /// The SendPort belonging to the helper isolate.
// Future<SendPort> _helperIsolateSendPort = () async {
//   // The helper isolate is going to send us back a SendPort, which we want to
//   // wait for.
//   final Completer<SendPort> completer = Completer<SendPort>();

//   // Receive port on the main isolate to receive messages from the helper.
//   // We receive two types of messages:
//   // 1. A port to send messages on.
//   // 2. Responses to requests we sent.
//   final ReceivePort receivePort = ReceivePort()
//     ..listen((dynamic data) {
//       if (data is SendPort) {
//         // The helper isolate sent us the port on which we can sent it requests.
//         completer.complete(data);
//         return;
//       }
//       if (data is _SumResponse) {
//         // The helper isolate sent us a response to a request we sent.
//         final Completer<int> completer = _sumRequests[data.id]!;
//         _sumRequests.remove(data.id);
//         completer.complete(data.result);
//         return;
//       }
//       throw UnsupportedError('Unsupported message type: ${data.runtimeType}');
//     });

//   // Start the helper isolate.
//   await Isolate.spawn((SendPort sendPort) async {
//     final ReceivePort helperReceivePort = ReceivePort()
//       ..listen((dynamic data) {
//         // On the helper isolate listen to requests and respond to them.
//         if (data is _SumRequest) {
//           final int result = _bindings.sum_long_running(data.a, data.b);
//           final _SumResponse response = _SumResponse(data.id, result);
//           sendPort.send(response);
//           return;
//         }
//         throw UnsupportedError('Unsupported message type: ${data.runtimeType}');
//       });

//     // Send the the port to the main isolate on which we can receive requests.
//     sendPort.send(helperReceivePort.sendPort);
//   }, receivePort.sendPort);

//   // Wait until the helper isolate has sent us back the SendPort on which we
//   // can start sending requests.
//   return completer.future;
// }();
