#include "srForm.h"

using namespace srProject1;
using namespace System::IO;

[STAThreadAttribute]

int main() {
	Application::Run(gcnew srForm());
	return 0;
}

System::Void srProject1::srForm::srForm_Load(System::Object ^ sender, System::EventArgs ^ e)
{
	String^ path = "C:\\Users\\chach\\Desktop\\samplecode.ini";
	//StreamReader^ sr = gcnew StreamReader(path);
	//String^ res = "";
	////res += sr->ReadToEnd();
	//while (true) {
	//	if (sr->EndOfStream) {
	//		res=res->Substring(0,res->Length-1);
	//		break;
	//	}
	//	else {
	//		res += sr->ReadLine()+"\r";
	//	}
	//}
	////MessageBox::Show(res);
	//sr->Close();
	String^ saveFile = "";
	StreamWriter^ sw = gcnew StreamWriter("C:\\Users\\chach\\Desktop\\samplecode.txt");
	StreamReader^ sr2 = gcnew StreamReader(path);
	while (true) {
		if (sr2->EndOfStream) {
			saveFile = saveFile->Substring(0, saveFile->Length - 1);
			break;
		}
		else {
			saveFile += sr2->ReadLine()+"\r";
		}
	}
	//sw->Write(res);
	sw->Write(saveFile);
	sw->Close();
	sr2->Close();
	return System::Void();
}
