// Reorder Worksheets.cpp : main project file.

#include "stdafx.h"
#include <iostream>
#include <msclr\marshal_cppstd.h>
using namespace System;
using namespace System::IO;
using namespace Microsoft::Office::Interop::Excel;
namespace Excel = Microsoft::Office::Interop::Excel;
#include <vector>
#include <string>
#include <map>
#include <string>
#include <iostream>
//
enum class eSheets { Facility = 1, Work_Breakdown_Structure, Drawings, Bill_of_Materials, Resources, TOGS, NSN, CostBook, Equipment };
int getIntFromSheetName(std::map<std::string, eSheets> &m, std::string sheetName);
std::string getSheetString(std::map<std::string, eSheets> &m, Excel::Workbook^ wb, int sheetNumber);
void moveSheet1Forward(Excel::Workbook^ wb, int i);
int main(array<System::String ^> ^args)
{
	String^ facPath = "C:\\Users\\CCrowe\\Documents\\AFCS Folder\\Facilities";
	DirectoryInfo^ di = gcnew DirectoryInfo(facPath);
	
	std::map<std::string, eSheets> sheetMap;
	sheetMap["Facility"] = eSheets::Facility;
	sheetMap["Work Breakdown Structure"] = eSheets::Work_Breakdown_Structure;
	sheetMap["Drawings"] = eSheets::Drawings;
	sheetMap["Bill of Materials"] = eSheets::Bill_of_Materials;
	sheetMap["Resources"] = eSheets::Resources;
	sheetMap["TOGS"] = eSheets::TOGS;
	sheetMap["NSN"] = eSheets::NSN;
	sheetMap["CostBook"] = eSheets::CostBook;
	sheetMap["Equipment"] = eSheets::Equipment;
	std::cout << sheetMap.size();
	for each(auto file in di->GetFiles("*")){
		Excel::Application^ xl = gcnew Excel::Application();
		xl->Visible = true;
		xl->DisplayAlerts = false;
		xl->AskToUpdateLinks = false;
		Excel::Workbook^ wb = xl->Workbooks->Open(file->FullName, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing);
		for (int j = 2; j <= wb->Sheets->Count; j++){
			auto keyName = getSheetString(sheetMap, wb, j);
			auto key = getIntFromSheetName(sheetMap, keyName);
				//sheetMap[msclr::interop::marshal_as<std::string>(static_cast<Excel::Worksheet^>(wb->Sheets[j])->Name->ToString()->Replace(" ", "_"))];
			int i = j - 1;
			auto aiString = getSheetString(sheetMap, wb, i);
				//msclr::interop::marshal_as<std::string>(static_cast<Excel::Worksheet^>(wb->Sheets[i])->Name->ToString()->Replace(" ", "_"));
			auto aiEnum = getIntFromSheetName(sheetMap, aiString);
			while (i > 0 && aiEnum > key){
				//a[i + 1] = a[i];
				if (i == 1){
					int p = 5;
				}
				moveSheet1Forward(wb, i);
				i = i - 1;
				if (i > 0){
					aiString = getSheetString(sheetMap, wb, i);
					aiEnum = getIntFromSheetName(sheetMap, aiString);
				}
			}
			//a[i + 1] = key;
			
			//keySheet->Move(Type::Missing, (System::Object^)static_cast<Excel::Worksheet^>(wb->Sheets[i]));
		}
		if (static_cast<_Worksheet^>(wb->Sheets[1])->Name == "Facility" &&
			static_cast<_Worksheet^>(wb->Sheets[2])->Name == "Work Breakdown Structure" &&
			static_cast<_Worksheet^>(wb->Sheets[3])->Name == "Drawings" &&
			static_cast<_Worksheet^>(wb->Sheets[4])->Name == "Bill of Materials" &&
			static_cast<_Worksheet^>(wb->Sheets[5])->Name == "Resources" &&
			static_cast<_Worksheet^>(wb->Sheets[6])->Name == "TOGS" &&
			static_cast<_Worksheet^>(wb->Sheets[7])->Name == "NSN" &&
			static_cast<_Worksheet^>(wb->Sheets[8])->Name == "CostBook" &&
			static_cast<_Worksheet^>(wb->Sheets[9])->Name == "Equipment"){
			static_cast<Excel::_Worksheet^>(wb->Sheets[1])->Activate();
			static_cast<Excel::_Worksheet^>(wb->Sheets[1])->Range["A1", Type::Missing]->Select();
			wb->Close((System::Object^)true, Type::Missing, Type::Missing);
		}
		else{
			std::cout << "Error!  Wrong Workbook Order! " << msclr::interop::marshal_as<std::string>(wb->Name) << "\n";
		}
		
		xl->Quit();
	}
    Console::WriteLine(L"Hello World");
    return 0;
}
int getIntFromSheetName(std::map<std::string, eSheets> &m, std::string sheetName){
	if (m.find(sheetName) != m.end()){
		auto found = m.find(sheetName);
		eSheets sheetEnum = found->second;//second gets value, first gets the key
		int enumVal = static_cast<int>(sheetEnum);
		return enumVal;
	}
	else{
		return -1;
	}
}
std::string getSheetString(std::map<std::string, eSheets> &m, Excel::Workbook^ wb, int sheetNumber){
	std::string sheetString = msclr::interop::marshal_as<std::string>(static_cast<Excel::Worksheet^>(wb->Sheets[sheetNumber])->Name->ToString());
	return sheetString;
}
void moveSheet1Forward(Excel::Workbook^ wb, int i){
	static_cast<Excel::Worksheet^>(wb->Sheets[i])->Move(Type::Missing, (System::Object^)static_cast<Excel::Worksheet^>(wb->Sheets[i + 1]));
}