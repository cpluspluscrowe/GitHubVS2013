// Minor Worksheet Fixes.cpp : main project file.

#include "stdafx.h"
#include <string>
#include <vector>
#include <msclr/marshal_cppstd.h>
#include <queue>
using namespace System;
using namespace System::IO;
using namespace Microsoft::Office::Interop::Excel;
namespace Excel = Microsoft::Office::Interop::Excel;
using namespace System::Threading;
#include "Mixed.h"
using namespace System::Threading;

ref class ManagedGlobals abstract sealed {
public:
	static std::queue<std::string>* files = new std::queue<std::string>[200];
	static System::Object^ obj = gcnew System::Object();
	static Random^ r = gcnew Random();
};


ref class ThreadFuncClass{
public:
	
	ThreadFuncClass(){
		fillFileQueue(ManagedGlobals::files, "C:\\Users\\CCrowe\\Documents\\AFCS Folder\\Facilities");
	}
	static void callFromThread(){
		std::string poppedName;
		Excel::Workbook^ wb;
		Excel::Application^ xl;
		try{
			//String^ facilityPath = "C:\\Users\\CCrowe\\Documents\\AFCS Folder\\Facilities";
			//DirectoryInfo^ di = gcnew DirectoryInfo(facilityPath);
			//for each(auto file in di->GetFiles("*")){
			String^ fileFullName;
			xl = gcnew Excel::Application();
			xl->Visible = true;
			xl->AskToUpdateLinks = false;
			xl->DisplayAlerts = false;
			xl->ScreenUpdating = true;
			xl->WindowState = Excel::XlWindowState::xlNormal;
			xl->Width = ManagedGlobals::r->Next(200, 1000);
			xl->Height = ManagedGlobals::r->Next(00, 1000);
			xl->Top = ManagedGlobals::r->Next(0, 100);
			xl->Left = ManagedGlobals::r->Next(-1000, 1000);
			while (!ManagedGlobals::files->empty()){
				Monitor::Enter(ManagedGlobals::obj);
				try
				{
					poppedName = ManagedGlobals::files->front();
					fileFullName = gcnew String(poppedName.c_str());
					ManagedGlobals::files->pop();//Pop once the changes to the file are finished
				}
				finally
				{
					Monitor::Exit(ManagedGlobals::obj);
				}
				//xl->WindowState = Excel::XlWindowState::xlMaximized;
				wb = xl->Workbooks->Open(fileFullName, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing);
				Excel::Worksheet^ ws = static_cast<Excel::Worksheet^>(wb->Sheets["Facility"]);
				bool changingFonts = false;
				ws->Range["B4", Type::Missing]->Formula = "='Bill of Materials'!T2";
				ws->Range["B5", Type::Missing]->Formula = "='Bill of Materials'!S2";
				ws->Range["B6", Type::Missing]->Formula = "='Bill of Materials'!R2";

				ws->Range["B8", Type::Missing]->Formula = "=Resources!J2";
				ws->Range["B9", Type::Missing]->Formula = "=Resources!K2 + Resources!L2 + Resources!O2 + Resources!P2";
				ws->Range["B10", Type::Missing]->Formula = "=Resources!N2";

				for (int i = 1; i <= 6; i++){
					ws = static_cast<Excel::Worksheet^>(wb->Sheets[i]);
					std::vector<std::string> styleNames;
					String^ wsName = ws->Name;
					static_cast<Excel::_Worksheet^>(ws)->Activate();
					for (int row = 1; row <= ws->UsedRange->Rows->Count; row++){
						Excel::Range^ rng = ws->Range["A" + row.ToString(), Type::Missing];
						if (rng->Value2 != nullptr){
							auto val = rng->Value2;
							if (rng->Value2->ToString() == "Key - Editable Columns are bold"){
								static_cast<Excel::Range^>(ws->Rows[i, Type::Missing])->AutoFit();
								Excel::Range^ mergedRange = ws->Range["A" + row.ToString(), "B" + row.ToString()];
								mergedRange->Font->Bold = false;
							}
							if (rng->Value2->ToString() == "26 10 00.00 0000 P128"){
								changingFonts = true;
								Excel::Range^ rngCell = ws->Range["A" + row.ToString(), Type::Missing];
								rngCell->Activate();
								String^ cellVal = ws->Range["B" + row.ToString(), Type::Missing]->Value2->ToString();
								Excel::Style^ styItem;
								for (int sty = 1; sty < wb->Styles->Count; sty++){
									styItem = wb->Styles->Item[sty];
									String^ styName = styItem->Name;
									if (styName == cellVal){
										styleNames.push_back(msclr::interop::marshal_as<std::string>(styName));
										break;
									}
								}
								if (cellVal == "Potential Problem/Missing Data"){
									rngCell->Style = styItem;//
								}
								else if (cellVal == "Designer Back Check Needed"){
									rngCell->Style = styItem; //"Designer Back Check Needed";
								}
								else if (cellVal == "Complete"){
									rngCell->Style = styItem; //"Complete";
								}
								else if (cellVal == "Untouched from JCMS"){
									rngCell->Style = styItem; //"Untouched from JCMS";
									rngCell->Borders[Excel::XlBordersIndex::xlEdgeBottom]->LineStyle = Excel::XlLineStyle::xlContinuous;
									rngCell->Borders[Excel::XlBordersIndex::xlEdgeBottom]->Weight = Excel::XlBorderWeight::xlMedium;
									if (styleNames[0].compare(std::string("Potential Problem/Missing Data")) == 0 &&
										styleNames[1].compare(std::string("Designer Back Check Needed")) == 0 &&
										styleNames[2].compare(std::string("Complete")) == 0 &&
										styleNames[3].compare(std::string("Untouched from JCMS")) == 0){
										//good
									}
									else{
										throw "There Was A Problem";
									}
								}
							}
						}
					}
				}
				if (changingFonts == false){
					throw "Did not Change Font";
				}
				String^ wbName = wb->Name;
				static_cast<Excel::_Worksheet^>(wb->Sheets[1])->Activate();
				static_cast<Excel::_Worksheet^>(wb->Sheets[1])->Range["A1", Type::Missing]->Select();
				wb->Close((System::Object^)true, Type::Missing, Type::Missing);
				writeOnFinish(msclr::interop::marshal_as<std::string>(wbName));
			}
			//}
			xl->Quit();
		}
		catch (int e){
			try{
				writeError(msclr::interop::marshal_as<std::string>(wb->Name));
				wb->Close((System::Object^)false, Type::Missing, Type::Missing);
				xl->Quit();
				ManagedGlobals::files->push(poppedName);
			}
			catch (int e2){
				writeError("Crud, threw error twice");
				ManagedGlobals::files->push(poppedName);
				writeError("Pushed back on");
				//do nothing, it must already be closed
			}
		}
	}
};


int main(array<System::String ^> ^args)
{
	ThreadFuncClass^ tfc = gcnew ThreadFuncClass();
	Thread^ t1 = gcnew Thread(gcnew ThreadStart(&ThreadFuncClass::callFromThread));
	t1->Start();
	Thread^ t2 = gcnew Thread(gcnew ThreadStart(&ThreadFuncClass::callFromThread));
	t2->Start();
	Thread^ t3 = gcnew Thread(gcnew ThreadStart(&ThreadFuncClass::callFromThread));
	t3->Start();

	Thread^ t4 = gcnew Thread(gcnew ThreadStart(&ThreadFuncClass::callFromThread));
	t4->Start();


	t1->Join();
	t2->Join();
	t3->Join();
	t4->Join();
	return 0;
}