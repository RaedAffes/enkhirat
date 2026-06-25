import 'dart:async';
import 'dart:io';

import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:excel/excel.dart' hide Border;
import 'package:path_provider/path_provider.dart';


void main() {
  runApp(const EnkhiratApp());
}

class EnkhiratApp extends StatelessWidget {
  const EnkhiratApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'enkhirat',       

      debugShowCheckedModeBanner: false,
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(
          seedColor: const Color(0xFF567357),
          brightness: Brightness.light,
        ),
        useMaterial3: true,
      ),
      home: const HomePage(),
    );
  }
}

class Person {
  final String serial;
  final String name;
  final String idCard;
  final String associationId;
  final String date;
  final String registrationNumber;
  final String year;

  Person({
    required this.serial,
    required this.name,
    required this.idCard,
    required this.associationId,
    required this.date,
    required this.registrationNumber,
    required this.year,
  });
}

class HomePage extends StatefulWidget {
  const HomePage({super.key});

  @override
  State<HomePage> createState() => _HomePageState();
}

class _HomePageState extends State<HomePage>
    with SingleTickerProviderStateMixin {
  List<Person> _allPeople = [];
  List<Person> _searchResults = [];
  final TextEditingController _searchController = TextEditingController();
  final FocusNode _searchFocus = FocusNode();
  final ScrollController _scrollController = ScrollController();
  Timer? _longPressTimer;
  String? _errorMessage;
  bool _fileLoaded = false;
  bool _hasSearched = false;

  String? _excelFilePath;
  Excel? _excel;

  late AnimationController _animController;
  late Animation<double> _fadeAnim;
  late Animation<Offset> _slideAnim;

  @override
  void initState() {
    super.initState();
    _loadExcelFromAssets();
    _animController = AnimationController(
      vsync: this,
      duration: const Duration(milliseconds: 500),
    );
    _fadeAnim =
        CurvedAnimation(parent: _animController, curve: Curves.easeOutCubic);
    _slideAnim = Tween<Offset>(
      begin: const Offset(0, 0.4),
      end: Offset.zero,
    ).animate(
        CurvedAnimation(parent: _animController, curve: Curves.easeOutBack));
  }

  Future<void> _saveExcel() async {
    if (_excel == null) return;

    final sheet = _excel!.sheets.values.first;
    final newPerson = _allPeople.last;

    sheet.appendRow(<CellValue?>[
      TextCellValue(newPerson.serial),
      TextCellValue(newPerson.name),
      TextCellValue(newPerson.idCard),
      TextCellValue(newPerson.associationId),
      TextCellValue(newPerson.date),
      TextCellValue(newPerson.registrationNumber),
      TextCellValue(newPerson.year),
    ]);

    final encoded = _excel!.encode();
    if (encoded == null) return;

    if (_excelFilePath != null) {
      await File(_excelFilePath!).writeAsBytes(encoded);
    }
  }

  Future<void> _openExcelFile() async {
    final path = _excelFilePath;
    if (path == null) {
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(
            content: const Text('لم يتم اختيار ملف بعد'),
            behavior: SnackBarBehavior.floating,
            backgroundColor: Colors.orange,
          ),
        );
      }
      return;
    }
    if (await File(path).exists()) {
      await Process.run('cmd', ['/c', 'start', '', path], runInShell: true);
    }
  }

  Future<void> _loadExcelFromAssets() async {
    try {
      final byteData = await rootBundle.load('assets/data.xlsx');
      final bytes = byteData.buffer.asUint8List(
        byteData.offsetInBytes,
        byteData.lengthInBytes,
      );

      _excel = Excel.decodeBytes(bytes);

      final sheet = _excel!.sheets.values.first;
      if (sheet.rows.isEmpty) {
        setState(() {
          _fileLoaded = true;
          _errorMessage = 'ملف Excel فارغ. لا توجد بيانات للقراءة.';
        });
        return;
      }

      final people = <Person>[];
      for (var i = 1; i < sheet.rows.length; i++) {
        final row = sheet.rows[i];
        final name = row[1]?.value?.toString().trim() ?? '';
        final idCard = row[2]?.value?.toString().trim() ?? '';

        if (name.isNotEmpty && idCard.isNotEmpty) {
          people.add(Person(
            serial: row[0]?.value?.toString().trim() ?? '',
            name: name,
            idCard: idCard,
            associationId: row[3]?.value?.toString().trim() ?? '',
            date: row[4]?.value?.toString().trim() ?? '',
            registrationNumber: row[5]?.value?.toString().trim() ?? '',
            year: row[6]?.value?.toString().trim() ?? '',
          ));
        }
      }

      final dir = await getApplicationDocumentsDirectory();
      final writablePath = '${dir.path}/data.xlsx';
      await File(writablePath).writeAsBytes(bytes);

      setState(() {
        _allPeople = people;
        _excelFilePath = writablePath;
        _fileLoaded = true;
        _errorMessage = null;
      });
    } catch (e) {
      setState(() {
        _errorMessage = 'خطأ في تحميل ملف البيانات:\n$e';
        _fileLoaded = true;
      });
    }
  }

  Future<void> _showAddPersonDialog() async {
    final nameController = TextEditingController();
    final idCardController = TextEditingController();
    final dateController = TextEditingController();
    final registrationNumberController = TextEditingController();
    final yearController = TextEditingController();
    final formKey = GlobalKey<FormState>();

    final result = await showDialog<bool>(
      context: context,
      barrierDismissible: false,
      builder: (ctx) {
        return Directionality(
          textDirection: TextDirection.rtl,
          child: AlertDialog(
            title: const Text(
              'إضافة منخرط جديد',
              textAlign: TextAlign.center,
              style: TextStyle(
                fontSize: 24,
                fontWeight: FontWeight.bold,
                color: Color(0xFF3D5A3D),
              ),
            ),
            content: Form(
              key: formKey,
              child: SingleChildScrollView(
                child: Column(
                  mainAxisSize: MainAxisSize.min,
                  children: [
                    _buildFormField(nameController, 'إسم المنخرط *', isRequired: true),
                    _buildFormField(idCardController, 'ب ت و'),
                    _buildFormField(dateController, 'التاريخ'),
                    _buildFormField(registrationNumberController, 'رقم الإنخراط'),
                    _buildFormField(yearController, 'السنة'),
                  ],
                ),
              ),
            ),
            actions: [
              TextButton(
                onPressed: () => Navigator.of(ctx).pop(false),
                child: const Text('إلغاء', style: TextStyle(fontSize: 18, color: Colors.grey)),
              ),
              ElevatedButton(
                onPressed: () {
                  if (formKey.currentState!.validate()) {
                    Navigator.of(ctx).pop(true);
                  }
                },
                style: ElevatedButton.styleFrom(
                  backgroundColor: const Color(0xFF567357),
                  foregroundColor: Colors.white,
                  padding: const EdgeInsets.symmetric(horizontal: 24, vertical: 12),
                ),
                child: const Text('حفظ', style: TextStyle(fontSize: 18)),
              ),
            ],
          ),
        );
      },
    );

    if (result == true) {
      final enteredYear = yearController.text.trim();
      final enteredName = nameController.text.trim();
      final enteredIdCard = idCardController.text.trim();
      final enteredDate = dateController.text.trim();
      final enteredRegNum = registrationNumberController.text.trim();
      final enteredYearVal = yearController.text.trim();

      // Auto-generate serial based on the same year
      final sameYearEntries = _allPeople.where((p) => p.year == enteredYear);
      String nextSerial;
      if (sameYearEntries.isNotEmpty) {
        final maxSerial = sameYearEntries
            .map((p) => int.tryParse(p.serial) ?? 0)
            .reduce((a, b) => a > b ? a : b);
        nextSerial = (maxSerial + 1).toString();
      } else {
        nextSerial = '1';
      }

      final newPerson = Person(
        serial: nextSerial,
        name: enteredName,
        idCard: enteredIdCard,
        associationId: '',
        date: enteredDate,
        registrationNumber: enteredRegNum,
        year: enteredYearVal,
      );

      _allPeople.add(newPerson);
      if (_searchController.text.isNotEmpty) {
        _search(_searchController.text);
      } else {
        setState(() {});
      }

      await _saveExcel();
    }

    nameController.dispose();
    idCardController.dispose();
    dateController.dispose();
    registrationNumberController.dispose();
    yearController.dispose();
  }

  Widget _buildFormField(TextEditingController controller, String label,
      {bool isRequired = false, TextInputType? keyboardType}) {
    return Padding(
      padding: const EdgeInsets.symmetric(vertical: 6),
      child: TextFormField(
        controller: controller,
        keyboardType: keyboardType,
        textAlign: TextAlign.right,
        textDirection: TextDirection.rtl,
        decoration: InputDecoration(
          labelText: label,
          labelStyle: const TextStyle(fontSize: 16),
          border: OutlineInputBorder(borderRadius: BorderRadius.circular(12)),
          filled: true,
          fillColor: Colors.grey[50],
        ),
        validator: isRequired
            ? (v) => (v == null || v.trim().isEmpty) ? 'هذا الحقل مطلوب' : null
            : null,
      ),
    );
  }

  void _search(String query) {
    final q = query.trim();
    final hadResults = _searchResults.isNotEmpty;

    setState(() {
      if (q.isEmpty) {
        _searchResults = [];
        _hasSearched = false;
      } else {
        final normalizedQ = _normalizeArabic(q);
        _searchResults = _allPeople.where((p) {
          return _normalizeArabic(p.name).contains(normalizedQ) ||
              p.idCard.contains(q);
        }).toList();
        _hasSearched = true;
      }
    });

    if (_searchResults.isNotEmpty && !hadResults) {
      _animController.forward(from: 0);
    } else if (_searchResults.isEmpty) {
      _animController.reset();
    } else {
      _animController.forward(from: 0);
    }
  }

  @override
  void dispose() {
    _longPressTimer?.cancel();
    _searchController.dispose();
    _searchFocus.dispose();
    _scrollController.dispose();
    _animController.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      body: Container(
        decoration: BoxDecoration(
          color: const Color(0xFF567357),
          image: DecorationImage(
            image: const AssetImage('assets/image/islamic.png'),
            fit: BoxFit.cover,
            colorFilter: ColorFilter.mode(
              const Color(0xFF567357).withValues(alpha: 0.85),
              BlendMode.srcOver,
            ),
          ),
        ),
        child: SafeArea(
          child: CallbackShortcuts(
            bindings: {
              SingleActivator(LogicalKeyboardKey.arrowUp): () =>
                  _scrollByKeyboard(-200),
              SingleActivator(LogicalKeyboardKey.arrowDown): () =>
                  _scrollByKeyboard(200),
            },
            child: Focus(
              autofocus: true,
                                          child: Column(
                                            children: [
                                              _buildTopBar(),
                                              const SizedBox(height: 16),
                                              Text(
                'قائمة المنخرطين لسنوات: 2024 - 2025 - 2026',
                style: TextStyle(
                  color: Colors.white.withValues(alpha: 0.95),
                  fontSize: 30,
                  fontWeight: FontWeight.bold,
                  letterSpacing: 1.2,
                  fontFamily: 'Traditional Arabic',
                  shadows: [
                    Shadow(
                      color: Colors.black.withValues(alpha: 0.3),
                      blurRadius: 8,
                      offset: const Offset(0, 2),
                    ),
                  ],
                ),
                textAlign: TextAlign.center,
              ),
              const SizedBox(height: 24),
              _buildSearchBar(),
              const SizedBox(height: 16),
              if (!_fileLoaded)
                Expanded(child: _buildLoadingView())
              else if (_errorMessage != null)
                Expanded(child: _buildErrorView())
              else
                Expanded(child: _buildResultsArea()),
            ],
          ),
        ),
        ),
        ),
      ),
    );
  }

  Widget _buildTopBar() {
    return Padding(
      padding: const EdgeInsets.fromLTRB(16, 16, 16, 0),
      child: Row(
        crossAxisAlignment: CrossAxisAlignment.center,
        children: [
          Expanded(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(
                  'جمعية المحافظة على القرآن الكريم',
                  style: TextStyle(
                    color: Colors.white.withValues(alpha: 0.95),
                    fontSize: 16,
                    fontWeight: FontWeight.bold,
                  ),
                ),
                Text(
                  'والأخلاق الفاضلة بصفاقس',
                  style: TextStyle(
                    color: Colors.white.withValues(alpha: 0.85),
                    fontSize: 14,
                    fontWeight: FontWeight.w600,
                  ),
                ),
                const SizedBox(height: 2),
                Text(
                  'فرع صفاقس المدينة',
                  style: TextStyle(
                    color: Colors.white.withValues(alpha: 0.7),
                    fontSize: 13,
                    fontWeight: FontWeight.w500,
                  ),
                ),
                Text(
                  'مركز البدر',
                  style: TextStyle(
                    color: Colors.white.withValues(alpha: 0.6),
                    fontSize: 12,
                    fontWeight: FontWeight.w500,
                  ),
                ),
              ],
            ),
          ),
          const SizedBox(width: 12),
          ClipOval(
            child: Image.asset(
              'assets/image/badr logo.png',
              width: 100,
              height: 100,
              fit: BoxFit.cover,
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildLoadingView() {
    return Center(
      child: Column(
        mainAxisAlignment: MainAxisAlignment.center,
        children: [
          const CircularProgressIndicator(color: Colors.white),
          const SizedBox(height: 20),
          Text(
            'جاري تحميل البيانات...',
            style: TextStyle(
              color: Colors.white.withValues(alpha: 0.8),
              fontSize: 22,
              fontWeight: FontWeight.bold,
            ),
            textAlign: TextAlign.center,
          ),
        ],
      ),
    );
  }

  Widget _buildErrorView() {
    return Center(
      child: Padding(
        padding: const EdgeInsets.all(24),
        child: Column(
          mainAxisAlignment: MainAxisAlignment.center,
          children: [
            Icon(Icons.error_outline,
                size: 64, color: Colors.white.withValues(alpha: 0.7)),
            const SizedBox(height: 16),
            Text(
              _errorMessage!,
              style: TextStyle(
                color: Colors.white.withValues(alpha: 0.9),
                fontSize: 16,
              ),
              textAlign: TextAlign.center,
            ),
            const SizedBox(height: 20),
            ElevatedButton.icon(
              onPressed: _loadExcelFromAssets,
              icon: const Icon(Icons.refresh),
              label: const Text('إعادة التحميل'),
              style: ElevatedButton.styleFrom(
                backgroundColor: Colors.white,
                foregroundColor: const Color(0xFF567357),
              ),
            ),
          ],
        ),
      ),
    );
  }

  Widget _buildResultsArea() {
    if (!_hasSearched) {
      return const SizedBox.shrink();
    }

    if (_searchResults.isEmpty) {
      return Center(
        child: Column(
          mainAxisSize: MainAxisSize.min,
          children: [
            Icon(Icons.search_off_rounded,
                size: 64, color: Colors.white.withValues(alpha: 0.5)),
            const SizedBox(height: 16),
            Text(
              'لا توجد نتائج للبحث',
              style: TextStyle(
                color: Colors.white.withValues(alpha: 0.6),
                fontSize: 28,
                fontWeight: FontWeight.bold,
              ),
            ),
          ],
        ),
      );
    }

    return FadeTransition(
      opacity: _fadeAnim,
      child: SlideTransition(
        position: _slideAnim,
        child: Column(
          children: [
            Expanded(
              child: RawScrollbar(
                controller: _scrollController,
                thumbVisibility: true,
                thickness: 14,
                radius: const Radius.circular(7),
                minThumbLength: 120,
                thumbColor: Colors.white.withValues(alpha: 0.7),
                child: ListView.builder(
                  controller: _scrollController,
                  padding: const EdgeInsets.symmetric(
                      horizontal: 16, vertical: 0),
                  itemCount: _searchResults.length,
                  itemBuilder: (context, index) {
                    return _buildPersonCard(_searchResults[index]);
                  },
                ),
              ),
            ),
            if (_searchResults.length > 3)
              Padding(
                padding: const EdgeInsets.only(bottom: 8),
                child: Row(
                  mainAxisAlignment: MainAxisAlignment.center,
                  children: [
                    _scrollButton(
                      Icons.keyboard_arrow_up,
                      onTap: () => _scrollBy(-100),
                      onLongPress: () => _scrollBy(-150),
                    ),
                    const SizedBox(width: 24),
                    _scrollButton(
                      Icons.keyboard_arrow_down,
                      onTap: () => _scrollBy(100),
                      onLongPress: () => _scrollBy(150),
                    ),
                  ],
                ),
              ),
          ],
        ),
      ),
    );
  }

  Widget _buildPersonCard(Person person) {
    return Card(
      margin: const EdgeInsets.only(bottom: 12),
      elevation: 6,
      shadowColor: Colors.black45,
      shape: RoundedRectangleBorder(
        borderRadius: BorderRadius.circular(16),
      ),
      child: Container(
        decoration: BoxDecoration(
          borderRadius: BorderRadius.circular(16),
          gradient: LinearGradient(
            begin: Alignment.topLeft,
            end: Alignment.bottomRight,
            colors: [
              Colors.white,
              const Color(0xFF567357).withValues(alpha: 0.07),
            ],
          ),
        ),
        child: Padding(
          padding: const EdgeInsets.all(20),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.end,
            children: [
              Row(
                textDirection: TextDirection.rtl,
                children: [
                  Icon(Icons.person,
                      color: const Color(0xFF567357), size: 32),
                  const SizedBox(width: 12),
                  Expanded(
                    child: Text(
                      person.name,
                      style: const TextStyle(
                        fontSize: 20,
                        fontWeight: FontWeight.bold,
                        color: Color(0xFF3D5A3D),
                      ),
                      textAlign: TextAlign.right,
                    ),
                  ),
                ],
              ),
                  const Divider(height: 32, thickness: 1),
              _infoRow('رقم بطاقة التعريف', person.idCard,
                  Icons.badge_outlined),
              if (person.registrationNumber.isNotEmpty)
                _infoRow(
                    'رقم الإنخراط', person.registrationNumber, Icons.numbers),
              if (person.date.isNotEmpty)
                _infoRow('تاريخ الإنخراط', person.date, Icons.calendar_today),
              if (person.year.isNotEmpty)
                _infoRow('سنة الإنخراط', person.year, Icons.event),
              if (person.associationId.isNotEmpty)
                _infoRow('معرّف الجمعيّة', person.associationId,
                    Icons.assignment_ind),
            ],
          ),
        ),
      ),
    );
  }

  Widget _infoRow(String label, String value, IconData icon) {
    return Padding(
      padding: const EdgeInsets.symmetric(vertical: 6),
      child: Row(
        textDirection: TextDirection.rtl,
        children: [
          Icon(icon, size: 22, color: const Color(0xFF567357)),
          const SizedBox(width: 12),
          SelectableText(
            '$label :',
            style: const TextStyle(
              fontSize: 18,
              fontWeight: FontWeight.bold,
              color: Color(0xFF3D3D3D),
            ),
          ),
          const SizedBox(width: 6),
          Expanded(
            child: SelectableText(
              value,
              style: const TextStyle(
                fontSize: 18,
                fontWeight: FontWeight.w500,
                color: Color(0xFF567357),
              ),
              textAlign: TextAlign.right,
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildSearchBar() {
    return Padding(
      padding: const EdgeInsets.symmetric(horizontal: 20),
      child: Container(
        decoration: BoxDecoration(
          borderRadius: BorderRadius.circular(30),
          boxShadow: [
            BoxShadow(
              color: Colors.black.withValues(alpha: 0.25),
              blurRadius: 20,
              offset: const Offset(0, 8),
            ),
          ],
        ),
        child: TextField(
          controller: _searchController,
          focusNode: _searchFocus,
          onChanged: _search,
          textAlign: TextAlign.right,
          textDirection: TextDirection.rtl,
          style: const TextStyle(
            fontSize: 22,
            fontWeight: FontWeight.w500,
            color: Color(0xFF3D3D3D),
          ),
          decoration: InputDecoration(
            hintText:
            'ابحث باستخدام الاسم أو رقم بطاقة التعريف',
            hintStyle: TextStyle(
              color: const Color(0xFF567357).withValues(alpha: 0.5),
              fontSize: 20,
              fontWeight: FontWeight.w400,
            ),
            filled: true,
            fillColor: Colors.white,
            prefixIcon: SizedBox(
              width: 56,
              child: Icon(
                Icons.search,
                color: const Color(0xFF567357).withValues(alpha: 0.6),
                size: 32,
              ),
            ),
            suffixIcon: SizedBox(
              width: 56,
              child: _searchController.text.isNotEmpty
                  ? IconButton(
                      icon: Icon(
                        Icons.clear,
                        color: const Color(0xFF567357)
                            .withValues(alpha: 0.6),
                      ),
                      onPressed: () {
                        _searchController.clear();
                        _search('');
                        _searchFocus.unfocus();
                      },
                    )
                  : null,
            ),
            border: OutlineInputBorder(
              borderRadius: BorderRadius.circular(30),
              borderSide: BorderSide.none,
            ),
            contentPadding: const EdgeInsets.symmetric(
              horizontal: 24,
              vertical: 22,
            ),
          ),
        ),
      ),
    );
  }

  void _scrollByKeyboard(double delta) {
    final target = _scrollController.offset + delta;
    _scrollController.animateTo(
      target.clamp(0, _scrollController.position.maxScrollExtent),
      duration: const Duration(milliseconds: 250),
      curve: Curves.easeOutCubic,
    );
  }

  String _normalizeArabic(String s) {
    return s
        .replaceAll('أ', 'ا')
        .replaceAll('إ', 'ا')
        .replaceAll('آ', 'ا')
        .replaceAll('ة', 'ه')
        .replaceAll('ؤ', 'و')
        .replaceAll('ئ', 'ي');
  }

  void _scrollBy(double delta) {
    final target = _scrollController.offset + delta;
    _scrollController.animateTo(
      target.clamp(0, _scrollController.position.maxScrollExtent),
      duration: const Duration(milliseconds: 400),
      curve: Curves.easeInOut,
    );
  }

  Widget _scrollButton(IconData icon,
      {required VoidCallback onTap, required VoidCallback onLongPress}) {
    return GestureDetector(
      onTap: onTap,
      onLongPressStart: (_) {
        onLongPress();
        _longPressTimer?.cancel();
        _longPressTimer = Timer.periodic(const Duration(milliseconds: 100), (_) {
          onLongPress();
        });
      },
      onLongPressEnd: (_) => _longPressTimer?.cancel(),
      onLongPressCancel: () => _longPressTimer?.cancel(),
      child: Container(
        decoration: BoxDecoration(
          color: Colors.white.withValues(alpha: 0.25),
          shape: BoxShape.circle,
          boxShadow: [
            BoxShadow(
              color: Colors.black.withValues(alpha: 0.3),
              blurRadius: 8,
              offset: const Offset(0, 2),
            ),
          ],
        ),
        padding: const EdgeInsets.all(12),
        child: Icon(icon, color: Colors.white, size: 36),
      ),
    );
  }
}
