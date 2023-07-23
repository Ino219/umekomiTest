#include "MyForm.h"

using namespace umekomiTest;

using namespace Microsoft::Office::Core;
using namespace Microsoft::Office::Interop::PowerPoint;

[STAThreadAttribute]

int main() {
	System::Windows::Forms::Application::Run(gcnew MyForm());
	return 0;
}

System::Void umekomiTest::MyForm::MyForm_Load(System::Object ^ sender, System::EventArgs ^ e)
{
	String^ path2 = "C:\\Users\\chach\\Desktop\\pptest2";
	
	System::Reflection::Assembly^ a = GetType()->Assembly;
	System::Resources::ResourceManager^ r = gcnew System::Resources::ResourceManager(String::Format(L"{0}.Resource", a->GetName()->Name), a);
	Bitmap^ bmp = static_cast<Bitmap^>(r->GetObject(L"56"));

	/*System::Reflection::Assembly^ myasm = System::Reflection::Assembly::GetExecutingAssembly();
	Bitmap^ bmp = gcnew Bitmap(myasm->GetManifestResourceStream("flower.png"));*/
	pictureBox1->Image = (Image^)(r->GetObject(L"56"));
	//パワポCOM参照のインスタンス
	Microsoft::Office::Interop::PowerPoint::Application^ apt = gcnew Microsoft::Office::Interop::PowerPoint::ApplicationClass();
	Presentations^ presen = apt->Presentations;

	Presentation^ presense1 = presen->Open("template_sample", MsoTriState::msoFalse, MsoTriState::msoFalse, MsoTriState::msoFalse);

	//PresentationClass^ presense1 = static_cast<Microsoft::Office::Interop::PowerPoint::PresentationClass^>(r->GetObject(L"templete_sample"));
	//String^ getpath=presense1->Path;


	//プレゼンテーション新規作成
	//Microsoft::Office::Interop::PowerPoint::Presentation^ presense1 = presen->Add(MsoTriState::msoFalse);
	try {
		//スライド追加

		//presense1->Slides->Add(1, Microsoft::Office::Interop::PowerPoint::PpSlideLayout::ppLayoutBlank);
		
		//presense1->SaveAs(path2, Microsoft::Office::Interop::PowerPoint::PpSaveAsFileType::ppSaveAsDefault, MsoTriState::msoTrue);
	}catch(Exception^ e) {
		MessageBox::Show(e->ToString());
	}
	finally{
		//閉じる
		//presense1->Close();
		apt->Quit();
	}
	//cli::array<System::String^>^ str = myasm->GetExecutingAssembly()->GetManifestResourceNames();
	
	//System::IO::Stream^ fs = myasm->GetExecutingAssembly()->GetManifestResourceStream("umekomiTest.MyForm.resources");
	//Bitmap^ bmp = gcnew Bitmap(myasm->GetManifestResourceStream("sample_pic.bmp"));
	//Image^ test = Image::FromStream(fs);
	//PictureBox1に表示
	

	return System::Void();
}
