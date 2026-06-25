import 'package:flutter_test/flutter_test.dart';
import 'package:enkhirat_app/main.dart';

void main() {
  testWidgets('App loads correctly', (WidgetTester tester) async {
    await tester.pumpWidget(const EnkhiratApp());
    expect(find.byType(EnkhiratApp), findsOneWidget);
  });
}
