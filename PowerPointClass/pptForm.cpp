#include "pptForm.h"

using namespace PowerPointClass;

using namespace Microsoft::Office::Core;
using namespace Microsoft::Office::Interop::PowerPoint;

[STAThreadAttribute]

int main() {
	System::Windows::Forms::Application::Run(gcnew pptForm());
	return 0;
}

System::Void PowerPointClass::pptForm::pptForm_Load(System::Object ^ sender, System::EventArgs ^ e)
{
	//インスタンスを定義
	Microsoft::Office::Interop::PowerPoint::Application^ apt = gcnew Microsoft::Office::Interop::PowerPoint::ApplicationClass();
	//起動ファイルを全取得
	apt=(Microsoft::Office::Interop::PowerPoint::Application^)System::Runtime::InteropServices::Marshal::GetActiveObject("PowerPoint.Application");
	Microsoft::Office::Interop::PowerPoint::Presentations^ presen = (Microsoft::Office::Interop::PowerPoint::Presentations^)apt->Presentations;
	
	String^ disposeFile;
	String^ disposePath;
	//起動しているファイルのフルパスのリストを取得
	System::Collections::Generic::List<String^>^ fullPathList = gcnew System::Collections::Generic::List<String^>;
	System::Collections::Generic::List<Microsoft::Office::Interop::PowerPoint::Presentation^>^ prList = gcnew System::Collections::Generic::List<Microsoft::Office::Interop::PowerPoint::Presentation^>;
	Microsoft::Office::Interop::PowerPoint::Presentations^ presen_copy = (Microsoft::Office::Interop::PowerPoint::Presentations^)apt->Presentations;

	
	//起動ファイルをリストに取得
	for each(Microsoft::Office::Interop::PowerPoint::Presentation^ pr in presen_copy) {
		String^ t = pr->FullName;
		//ディレクトリパスを取得
		String^ p = pr->Path;
		if (t->IndexOf("/") >= 0 || t->IndexOf("\\") >= 0) {
			try {
				//フルパスを取得
				disposeFile = System::IO::Path::GetFileName(t);
				fullPathList->Add(p+"\\"+disposeFile);
				//保存して閉じる
				//pr->Save();
				prList->Add(pr);
				pr->Save();
				pr->Close();
				System::Runtime::InteropServices::Marshal::ReleaseComObject(pr);
			}
			catch (Exception^ e) {
				MessageBox::Show(e->ToString());
			}
		}
		else {
			MessageBox::Show(t);
		}
		
	}
	//disposePath = "C:\\Users\\chach\\Desktop\\templete1.pptx";
	/*System::IO::FileStream^ fs = gcnew System::IO::FileStream(disposePath, System::IO::FileMode::Append);
	fs->Close();*/
	

	//プレゼンテーション新規作成
	Microsoft::Office::Interop::PowerPoint::Presentation^ presense1 = presen->Open("C:\\Users\\chach\\Desktop\\sam.pptx", MsoTriState::msoFalse, MsoTriState::msoFalse, MsoTriState::msoTrue);
	//Microsoft::Office::Interop::PowerPoint::Presentation^ presense2 = presen->Open("C:\\Users\\chach\\Desktop\\pptest2.pptx", MsoTriState::msoFalse, MsoTriState::msoFalse, MsoTriState::msoTrue);
	//MessageBox::Show(apt->Presentations->Count.ToString()+":Close前");
	////閉じる
	presense1->Save();
	presense1->Close();

	//MessageBox::Show(apt->Presentations->Count.ToString()+"Close後");
	System::Runtime::InteropServices::Marshal::ReleaseComObject(presense1);
	for each (Microsoft::Office::Interop::PowerPoint::Presentation^ var in prList)
	{
		System::Runtime::InteropServices::Marshal::ReleaseComObject(var);
	}
	prList->Clear();
	System::Runtime::InteropServices::Marshal::ReleaseComObject(presen);
	System::Runtime::InteropServices::Marshal::ReleaseComObject(presen_copy);

	apt->Quit();
	System::Runtime::InteropServices::Marshal::ReleaseComObject(apt);
	//System::Runtime::InteropServices::Marshal::FinalReleaseComObject(apt);



	//インスタンスを定義
	Microsoft::Office::Interop::PowerPoint::Application^ apt2 = gcnew Microsoft::Office::Interop::PowerPoint::ApplicationClass();
	Microsoft::Office::Interop::PowerPoint::Presentations^ presen2 = apt2->Presentations;
	for each (String^ var in fullPathList)
	{
		MessageBox::Show(var);
		Microsoft::Office::Interop::PowerPoint::Presentation^ presense1 = presen2->Open(var, MsoTriState::msoFalse, MsoTriState::msoFalse, MsoTriState::msoTrue);
	
	}
	Microsoft::Office::Interop::PowerPoint::Presentation^ presense2 = presen2->Open("C:\\Users\\chach\\Desktop\\sam.pptx", MsoTriState::msoFalse, MsoTriState::msoFalse, MsoTriState::msoTrue);

	
	return System::Void();
}
