using namespace System;
using namespace Microsoft::Office::Interop::Excel;
int main(){
	Microsoft::Office::Interop::Excel::Application^ exApp = gcnew Microsoft::Office::Interop::Excel::ApplicationClass();
	String^ filename = "C:\\Users\\CCrowe\\Documents\\AFCS Folder\\Old_Scope_Facilities\\1411300BA - Entry Control Point Vehicle Inspection Building, 18.5X15.5 METERS.xlsm";
	Workbook^ wb = exApp->Workbooks->Open(filename, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing);
	Worksheet^ exWs = safe_cast<Worksheet^>(exApp->ActiveSheet);
	int row = 2;
	int col = 1;
	String^ tmp = ((Microsoft::Office::Interop::Excel::Range^)exWs->Cells[(System::Object^)row, (System::Object^)col])->Value2->ToString();
	Console::WriteLine(tmp);
	Console::ReadLine();
	return 0;
}