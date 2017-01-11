// Excel Delete Repeats.cpp : main project file.
using namespace Microsoft::Office::Interop::Excel;
#include "stdafx.h"
namespace Excel = Microsoft::Office::Interop::Excel;
using namespace System;
using namespace System::Collections::Generic;
#using <System.Core.dll>

int main(array<System::String ^> ^args)
{
	Excel::Application^ xl = gcnew Excel::Application;
	xl->Visible = true;
	String^ file = "C:\\Users\\CCrowe\\Documents\\AFCS Folder\\New Scope Workbook Errors.xlsx";
	Excel::Workbook^ wb = xl->Workbooks->Open(file, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing);
	Excel::Worksheet^ exWs = safe_cast<Excel::Worksheet^>(xl->ActiveSheet);
	HashSet<String^>^ hashes = gcnew HashSet<String^>();
	for (Int32 i = 1; i <= exWs->UsedRange->Rows->Count; i++){
		String^ cellValue = exWs->Range["A" + Int32(i).ToString(), "A" + Int32(i).ToString()]->Value2->ToString();
		if (cellValue->Length == 16){
			if (cellValue != ""){
				exWs->Range["A" + Int32(i).ToString(), "A" + Int32(i).ToString()]->Delete(Excel::XlDeleteShiftDirection::xlShiftUp);
				i -= 1;
				hashes->Add(cellValue);
			}
		}
	}
	for each(String^ val in hashes){
		Console::WriteLine(val);
	}
	Console::ReadLine();
    return 0;
}
