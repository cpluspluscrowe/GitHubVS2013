// Find Bad Newly Created Facility Sheets.cpp : main project file.

#include "stdafx.h";
#include <msclr\marshal_cppstd.h>;
using namespace System;
using namespace System::IO;
using namespace Microsoft::Office::Interop::Excel;
#include <iostream>;
#include <fstream>;
using namespace std;
using namespace System::IO;
namespace Excel = Microsoft::Office::Interop::Excel;

int findKey(Excel::Worksheet^ ws);//can't call the funciton until you include the header somewhere
int findUsedRange(Excel::Worksheet^ ws);
void writeToFile(string errorString);

ref class ManagedGlobals abstract sealed {
public:
	static Excel::Application^ xl;
};

int main(array<System::String ^> ^args)
{
	string s = "C:\\Users\\CCrowe\\Documents\\AFCS Folder\\Facilities";
	String^ managed = gcnew String(s.c_str());
	DirectoryInfo^ di = gcnew DirectoryInfo(managed);
	for each(auto file in di->GetFiles("*")){
		ManagedGlobals::xl = gcnew Excel::Application();
		ManagedGlobals::xl->Visible = true;
		Excel::Workbook^ wb = ManagedGlobals::xl->Workbooks->Open(file->FullName, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing);
		for (int i = 1; i <= System::Math::Min(wb->Sheets->Count,6); i++){
			Excel::Worksheet^ ws = static_cast<Excel::Worksheet^>(wb->Sheets[i]);
			//static_cast<Microsoft::Office::Interop::Excel::_Worksheet^>(ws)->Activate();
			if (i == 1){
				ws->Range["B4", "B4"]->Formula = "='Bill of Materials'!T2";
				ws->Range["B5", "B5"]->Formula = "='Bill of Materials'!S2";
				ws->Range["B6", "B6"]->Formula = "='Bill of Materials'!R2";
				ws->Range["B8", "B8"]->Formula = "=Resources!J2";
				ws->Range["B9", "B9"]->Formula = "=Resources!K2 + Resources!L2 + Resources!O2 + Resources!P2";
				ws->Range["B10", "B10"]->Formula = "=Resources!N2";
			}
			else{
				int keyRow = findKey(ws);
				int wsUsedRange = findUsedRange(ws);
				if (wsUsedRange > keyRow + 4){
					string errorString = msclr::interop::marshal_as<std::string>(wb->Name) + " : " + msclr::interop::marshal_as<std::string>(ws->Name) + " " +
						msclr::interop::marshal_as<std::string>(keyRow.ToString()) + " " + msclr::interop::marshal_as<std::string>(wsUsedRange.ToString()) + "\n";
					writeToFile(errorString);
				}
			}
		}
		wb->Close((System::Object^)false,Type::Missing,Type::Missing);
		string fileName = msclr::interop::marshal_as<std::string>(file->FullName);
		std::remove(fileName.c_str());
		ManagedGlobals::xl->Quit();
	}
	Console::ReadLine();
    return 0;
}
int findKey(Excel::Worksheet^ ws){
	String^ wsName = ws->Name;
	auto cellValu2 = ws->Range["A1", "A1"]->Value2;
	for (int i = 1; i <= ws->UsedRange->Rows->Count; i++){
		Range^ rng = ws->Range["A" + i.ToString(), "B" + i.ToString()];
		bool isMerged = (bool)rng->MergeCells;
		if (isMerged){
			 //auto singleCell = ((Excel::Range^)ws->Cells[(System::Object^)i, (System::Object^)2]);
			if (ws->Range[((Excel::Range^)ws->Cells[(System::Object^)i, (System::Object^)1]), ((Excel::Range^)ws->Cells[(System::Object^)i, (System::Object^)1])]->Value2 != nullptr){
				auto cellValue = ws->Range[((Excel::Range^)ws->Cells[(System::Object^)i, (System::Object^)1]), ((Excel::Range^)ws->Cells[(System::Object^)i, (System::Object^)1])]->Value2->ToString();
				if (cellValue != nullptr){
					if (cellValue == "Key - Editable Columns are bold"){
						return i;
					}
				}
			}
		}
	}
	try{
		throw 0;
	}
	catch (int e){
		cout << " An Exception has occurred finding they key!\n";
	}
	return 0;
}
int findUsedRange(Excel::Worksheet^ ws){
	int lastRow = -1;
	for (int i = 1; i <= ws->UsedRange->Rows->Count; i++){
		Range^ rng = ws->Range["A" + i.ToString(), "Z" + i.ToString()];
		try{
			if (ManagedGlobals::xl->WorksheetFunction->CountA(rng, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing) > 0){
				lastRow = i;
			}
		}
		catch (int e){

		}
	}
	return lastRow;
}
#pragma unmanaged
void writeToFile(string errorString){
	ofstream myfile;
	myfile.open("Excel Formatting Errors.txt", std::ios_base::app);
	myfile << errorString;
	myfile.close();
}