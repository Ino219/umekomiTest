#include "scrollForm.h"

using namespace fileScroll;
using namespace System::Diagnostics;
using namespace windowCapture1;
using namespace System::Threading::Tasks;
using namespace System::Collections::Generic;

[STAThreadAttribute]

int main() {
	Application::Run(gcnew scrollForm());
	return 0;
		
}

System::Void fileScroll::scrollForm::scrollForm_Load(System::Object ^ sender, System::EventArgs ^ e)
{
	//現在、開いているフォルダのリストを取得
	List<String^>^ temp = windowCapture1::Form1::getOpenFolderList();
	//比較対象のファイルを定義
	String^ FilePath = "C:\\Users\\chach\\Desktop\\testFolder\\test.txt";
	//ファイルのフルパスを\\文字で区切る
	array<String^>^ checkFolderArray = FilePath->Split('\\');
	//パスを抜いたファイル名を取得
	String^ FileSinglePath = checkFolderArray[checkFolderArray->Length - 1];
	//そのファイル名と区切り文字1文字分を抜いたフォルダパスを取得
	String^ FolderPath = FilePath->Substring(0, (FilePath->Length - (FileSinglePath->Length+1)));
	//\\という区切り文字でパスを分割し、配列化する
	array<String^>^ checkFolderPathArray = FolderPath->Split('\\');
	//フォルダ名を取得
	String^ FolderSinglePath = checkFolderPathArray[checkFolderPathArray->Length - 1];
	//現在開いているフォルダと開こうとしているファイルのフォルダが一致しているかどうかを判定
	bool matchFlg = false;

	for each (String^ var in temp)
	{
		//一致するフォルダがあれば、起動中フォルダに探しているフォルダがあるのでフラグをtrueにする
		if (var == FolderPath) {
			matchFlg = true;
		}
	}

	//フォルダ名を指定して、アクティブウィンドウ化する
	//IntPtr^ ip = windowCapture1::Form1::FindWindowByName("testFolder");
	//windowCapture1::Form1::SetForeGroundByIntPtr((IntPtr)ip);

	//ファイルパスを指定して、そのファイルが含まれるフォルダを最前面に開く
	//既に開いている場合は再利用して、最前面に移動する
	windowCapture1::Form1::setForeWindowSelectFile(FilePath);


	//try {
	//	指定したファイルパスを選択した状態でフォルダを開き、最前面にある状態にする
	//	ただし、既に開いている場合は二重に開いてしまう
	//	Process::Start("explorer.exe", "/select, ""C:\\Users\\chach\\Desktop\\testFolder\\test.txt");
	//	//エクスプローラが起動するまで待機
	//	System::Threading::Thread::Sleep(1500);
	//}
	//catch (Exception^ e) {
	//	MessageBox::Show(e->ToString());
	//}
	//sleepではなくdelayで待機処理に挑戦したが、上手くいかず
	////Task^t = windowCapture1::Form1::AsyncFunction();
	////t->Wait();
	//現在のアクティブウィンドウを取得して、画像として保存する
	//String^ Path = windowCapture1::Form1::GetGazouPath();
	return System::Void();
}






