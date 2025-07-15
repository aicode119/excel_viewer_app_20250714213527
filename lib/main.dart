import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel/excel.dart';
import 'package:shared_preferences/shared_preferences.dart';
import 'package:permission_handler/permission_handler.dart';
import 'dart:io';
import 'dart:typed_data';

void main() {
  runApp(const ExcelViewerApp());
}

class ExcelViewerApp extends StatelessWidget {
  const ExcelViewerApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Excel Viewer',
      theme: ThemeData(
        primarySwatch: Colors.blue,
        useMaterial3: true,
      ),
      darkTheme: ThemeData.dark(
        useMaterial3: true,
      ),
      themeMode: ThemeMode.system,
      home: const ExcelViewerHome(),
    );
  }
}

class ExcelViewerHome extends StatefulWidget {
  const ExcelViewerHome({super.key});

  @override
  State<ExcelViewerHome> createState() => _ExcelViewerHomeState();
}

class _ExcelViewerHomeState extends State<ExcelViewerHome> {
  Excel? _excel;
  String? _fileName;
  int? _fileSize;
  List<String> _sheetNames = [];
  String? _selectedSheet;
  List<List<Data?>> _currentSheetData = [];
  List<List<Data?>> _filteredData = [];
  
  // Search controllers
  final TextEditingController _searchController1 = TextEditingController();
  final TextEditingController _searchController2 = TextEditingController();
  
  // Search options
  String _searchType1 = 'contains';
  String _searchType2 = 'contains';
  bool _searchAllSheets = false;
  bool _ignoreCase = true;
  
  // Pagination
  int _currentPage = 0;
  final int _rowsPerPage = 50;
  
  // Statistics
  int _totalRows = 0;
  int _totalColumns = 0;
  int _searchResults = 0;

  @override
  void initState() {
    super.initState();
    _loadLastFile();
  }

  Future<void> _loadLastFile() async {
    final prefs = await SharedPreferences.getInstance();
    final lastFilePath = prefs.getString('last_file_path');
    
    if (lastFilePath != null && File(lastFilePath).existsSync()) {
      await _loadExcelFile(lastFilePath);
    }
  }

  Future<void> _saveLastFile(String filePath) async {
    final prefs = await SharedPreferences.getInstance();
    await prefs.setString('last_file_path', filePath);
  }

  Future<void> _pickFile() async {
    if (await Permission.manageExternalStorage.request().isGranted) {
      FilePickerResult? result = await FilePicker.platform.pickFiles(
        type: FileType.custom,
        allowedExtensions: ["xlsx", "xls"],
      );

      if (result != null) {
        final file = result.files.first;
        if (file.path != null) {
          await _loadExcelFile(file.path!);
          await _saveLastFile(file.path!);
        }
      }
    } else {
      openAppSettings(); // فتح الإعدادات يدوياً إذا المستخدم رفض الإذن
      _showSnackBar("يرجى تفعيل إذن الوصول من الإعدادات");
    }
  }

  Future<void> _loadExcelFile(String filePath) async {
    try {
      final file = File(filePath);
      final bytes = await file.readAsBytes();
      
      setState(() {
        _excel = Excel.decodeBytes(bytes);
        _fileName = file.path.split('/').last;
        _fileSize = bytes.length;
        _sheetNames = _excel!.tables.keys.toList();
        _selectedSheet = _sheetNames.isNotEmpty ? _sheetNames.first : null;
      });
      
      if (_selectedSheet != null) {
        _loadSheetData(_selectedSheet!);
      }
      
      _showSnackBar('تم تحميل الملف بنجاح');
    } catch (e) {
      _showSnackBar('خطأ في تحميل الملف: $e');
    }
  }

  void _loadSheetData(String sheetName) {
    if (_excel == null) return;
    
    final table = _excel!.tables[sheetName];
    if (table == null) return;
    
    setState(() {
      _selectedSheet = sheetName;
      _currentSheetData = table.rows;
      _filteredData = List.from(_currentSheetData);
      _totalRows = _currentSheetData.length;
      _totalColumns = _currentSheetData.isNotEmpty ? _currentSheetData.first.length : 0;
      _searchResults = _filteredData.length;
      _currentPage = 0;
    });
  }

  void _performSearch() {
    if (_excel == null) return;
    
    final query1 = _searchController1.text.trim();
    final query2 = _searchController2.text.trim();
    
    if (query1.isEmpty && query2.isEmpty) {
      setState(() {
        _filteredData = List.from(_currentSheetData);
        _searchResults = _filteredData.length;
        _currentPage = 0;
      });
      return;
    }
    
    List<List<Data?>> results = [];
    
    if (_searchAllSheets) {
      // Search in all sheets
      for (String sheetName in _sheetNames) {
        final table = _excel!.tables[sheetName];
        if (table != null) {
          results.addAll(_filterRows(table.rows, query1, query2));
        }
      }
    } else {
      // Search in current sheet only
      results = _filterRows(_currentSheetData, query1, query2);
    }
    
    setState(() {
      _filteredData = results;
      _searchResults = results.length;
      _currentPage = 0;
    });
  }

  List<List<Data?>> _filterRows(List<List<Data?>> rows, String query1, String query2) {
    return rows.where((row) {
      bool match1 = query1.isEmpty || _matchesQuery(row, query1, _searchType1);
      bool match2 = query2.isEmpty || _matchesQuery(row, query2, _searchType2);
      return match1 && match2;
    }).toList();
  }

  bool _matchesQuery(List<Data?> row, String query, String searchType) {
    final searchQuery = _ignoreCase ? query.toLowerCase() : query;
    
    for (Data? cell in row) {
      if (cell?.value == null) continue;
      
      String cellValue = cell!.value.toString();
      if (_ignoreCase) cellValue = cellValue.toLowerCase();
      
      bool matches = false;
      switch (searchType) {
        case 'contains':
          matches = cellValue.contains(searchQuery);
          break;
        case 'exact':
          matches = cellValue == searchQuery;
          break;
        case 'startsWith':
          matches = cellValue.startsWith(searchQuery);
          break;
        case 'endsWith':
          matches = cellValue.endsWith(searchQuery);
          break;
        case 'wildcard':
          // Simple wildcard implementation
          final pattern = searchQuery.replaceAll('*', '.*');
          matches = RegExp(pattern).hasMatch(cellValue);
          break;
      }
      
      if (matches) return true;
    }
    
    return false;
  }

  void _showSnackBar(String message) {
    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(content: Text(message)),
    );
  }

  String _formatFileSize(int bytes) {
    if (bytes < 1024) return '$bytes B';
    if (bytes < 1024 * 1024) return '${(bytes / 1024).toStringAsFixed(1)} KB';
    return '${(bytes / (1024 * 1024)).toStringAsFixed(1)} MB';
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('Excel Viewer'),
        backgroundColor: Theme.of(context).colorScheme.inversePrimary,
      ),
      body: Column(
        children: [
          // File selection section
          Padding(
            padding: const EdgeInsets.all(16.0),
            child: Column(
              children: [
                ElevatedButton.icon(
                  onPressed: _pickFile,
                  icon: const Icon(Icons.file_open),
                  label: const Text('اختيار ملف Excel'),
                  style: ElevatedButton.styleFrom(
                    minimumSize: const Size(double.infinity, 50),
                  ),
                ),
                
                if (_fileName != null) ...[
                  const SizedBox(height: 16),
                  Card(
                    child: Padding(
                      padding: const EdgeInsets.all(16.0),
                      child: Column(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          Text('اسم الملف: $_fileName', style: const TextStyle(fontWeight: FontWeight.bold)),
                          Text('الحجم: ${_formatFileSize(_fileSize!)}'),
                          Text('عدد الشيتات: ${_sheetNames.length}'),
                          Text('عدد الأعمدة: $_totalColumns'),
                          Text('عدد الصفوف: $_totalRows'),
                          if (_searchResults != _totalRows)
                            Text('نتائج البحث: $_searchResults', style: const TextStyle(color: Colors.blue)),
                        ],
                      ),
                    ),
                  ),
                ],
              ],
            ),
          ),
          
          // Search section
          if (_excel != null) ...[
            Padding(
              padding: const EdgeInsets.symmetric(horizontal: 16.0),
              child: Card(
                child: Padding(
                  padding: const EdgeInsets.all(16.0),
                  child: Column(
                    children: [
                      // First search bar
                      Row(
                        children: [
                          Expanded(
                            flex: 3,
                            child: TextField(
                              controller: _searchController1,
                              decoration: const InputDecoration(
                                labelText: 'البحث الأول',
                                border: OutlineInputBorder(),
                                prefixIcon: Icon(Icons.search),
                              ),
                              onChanged: (_) => _performSearch(),
                            ),
                          ),
                          const SizedBox(width: 8),
                          Expanded(
                            flex: 2,
                            child: DropdownButtonFormField<String>(
                              value: _searchType1,
                              decoration: const InputDecoration(
                                border: OutlineInputBorder(),
                              ),
                              items: const [
                                DropdownMenuItem(value: 'contains', child: Text('يحتوي على')),
                                DropdownMenuItem(value: 'exact', child: Text('مطابقة تامة')),
                                DropdownMenuItem(value: 'startsWith', child: Text('يبدأ بـ')),
                                DropdownMenuItem(value: 'endsWith', child: Text('ينتهي بـ')),
                                DropdownMenuItem(value: 'wildcard', child: Text('بحث بالنجمة (*)')),
                              ],
                              onChanged: (value) {
                                setState(() {
                                  _searchType1 = value!;
                                });
                                _performSearch();
                              },
                            ),
                          ),
                        ],
                      ),
                      
                      const SizedBox(height: 12),
                      
                      // Second search bar
                      Row(
                        children: [
                          Expanded(
                            flex: 3,
                            child: TextField(
                              controller: _searchController2,
                              decoration: const InputDecoration(
                                labelText: 'البحث الثاني (اختياري)',
                                border: OutlineInputBorder(),
                                prefixIcon: Icon(Icons.search),
                              ),
                              onChanged: (_) => _performSearch(),
                            ),
                          ),
                          const SizedBox(width: 8),
                          Expanded(
                            flex: 2,
                            child: DropdownButtonFormField<String>(
                              value: _searchType2,
                              decoration: const InputDecoration(
                                border: OutlineInputBorder(),
                              ),
                              items: const [
                                DropdownMenuItem(value: 'contains', child: Text('يحتوي على')),
                                DropdownMenuItem(value: 'exact', child: Text('مطابقة تامة')),
                                DropdownMenuItem(value: 'startsWith', child: Text('يبدأ بـ')),
                                DropdownMenuItem(value: 'endsWith', child: Text('ينتهي بـ')),
                                DropdownMenuItem(value: 'wildcard', child: Text('بحث بالنجمة (*)')),
                              ],
                              onChanged: (value) {
                                setState(() {
                                  _searchType2 = value!;
                                });
                                _performSearch();
                              },
                            ),
                          ),
                        ],
                      ),
                      
                      const SizedBox(height: 12),
                      
                      // Search options
                      Row(
                        children: [
                          Expanded(
                            child: CheckboxListTile(
                              title: const Text('البحث في كافة المصنفات'),
                              value: _searchAllSheets,
                              onChanged: (value) {
                                setState(() {
                                  _searchAllSheets = value!;
                                });
                                _performSearch();
                              },
                            ),
                          ),
                          Expanded(
                            child: CheckboxListTile(
                              title: const Text('تجاهل حالة الأحرف'),
                              value: _ignoreCase,
                              onChanged: (value) {
                                setState(() {
                                  _ignoreCase = value!;
                                });
                                _performSearch();
                              },
                            ),
                          ),
                        ],
                      ),
                    ],
                  ),
                ),
              ),
            ),
            
            // Sheet selector
            if (_sheetNames.isNotEmpty)
              Padding(
                padding: const EdgeInsets.symmetric(horizontal: 16.0),
                child: DropdownButtonFormField<String>(
                  value: _selectedSheet,
                  decoration: const InputDecoration(
                    labelText: 'اختيار الشيت',
                    border: OutlineInputBorder(),
                  ),
                  items: _sheetNames.map((sheet) => 
                    DropdownMenuItem(value: sheet, child: Text(sheet))
                  ).toList(),
                  onChanged: (value) {
                    if (value != null) {
                      _loadSheetData(value);
                    }
                  },
                ),
              ),
            
            const SizedBox(height: 16),
            
            // Data table
            Expanded(
              child: _buildDataTable(),
            ),
            
            // Pagination
            if (_filteredData.length > _rowsPerPage)
              _buildPagination(),
          ],
        ],
      ),
    );
  }

  Widget _buildDataTable() {
    if (_filteredData.isEmpty) {
      return const Center(
        child: Text('لا توجد بيانات للعرض'),
      );
    }
    
    final startIndex = _currentPage * _rowsPerPage;
    final endIndex = (startIndex + _rowsPerPage).clamp(0, _filteredData.length);
    final pageData = _filteredData.sublist(startIndex, endIndex);
    
    return SingleChildScrollView(
      scrollDirection: Axis.horizontal,
      child: SingleChildScrollView(
        child: DataTable(
          columns: List.generate(
            _totalColumns,
            (index) => DataColumn(
              label: Text('العمود ${index + 1}'),
            ),
          ),
          rows: pageData.map((row) {
            return DataRow(
              cells: List.generate(
                _totalColumns,
                (index) => DataCell(
                  Text(
                    index < row.length && row[index]?.value != null
                        ? row[index]!.value.toString()
                        : '',
                  ),
                ),
              ),
            );
          }).toList(),
        ),
      ),
    );
  }

  Widget _buildPagination() {
    final totalPages = (_filteredData.length / _rowsPerPage).ceil();
    
    return Padding(
      padding: const EdgeInsets.all(16.0),
      child: Row(
        mainAxisAlignment: MainAxisAlignment.center,
        children: [
          IconButton(
            onPressed: _currentPage > 0 ? () {
              setState(() {
                _currentPage--;
              });
            } : null,
            icon: const Icon(Icons.chevron_left),
          ),
          
          ...List.generate(
            totalPages.clamp(0, 5),
            (index) {
              final pageIndex = _currentPage < 3 
                  ? index 
                  : (_currentPage + index - 2).clamp(0, totalPages - 1);
              
              return Padding(
                padding: const EdgeInsets.symmetric(horizontal: 4.0),
                child: ElevatedButton(
                  onPressed: () {
                    setState(() {
                      _currentPage = pageIndex;
                    });
                  },
                  style: ElevatedButton.styleFrom(
                    backgroundColor: _currentPage == pageIndex 
                        ? Theme.of(context).primaryColor 
                        : null,
                  ),
                  child: Text('${pageIndex + 1}'),
                ),
              );
            },
          ),
          
          IconButton(
            onPressed: _currentPage < totalPages - 1 ? () {
              setState(() {
                _currentPage++;
              });
            } : null,
            icon: const Icon(Icons.chevron_right),
          ),
        ],
      ),
    );
  }

  @override
  void dispose() {
    _searchController1.dispose();
    _searchController2.dispose();
    super.dispose();
  }
}



