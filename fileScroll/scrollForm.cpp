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
	//���݁A�J���Ă���t�H���_�̃��X�g���擾
	List<String^>^ temp = windowCapture1::Form1::getOpenFolderList();
	//��r�Ώۂ̃t�@�C�����`
	String^ FilePath = "C:\\Users\\chach\\Desktop\\testFolder\\test.txt";
	//�t�@�C���̃t���p�X��\\�����ŋ�؂�
	array<String^>^ checkFolderArray = FilePath->Split('\\');
	//�p�X�𔲂����t�@�C�������擾
	String^ FileSinglePath = checkFolderArray[checkFolderArray->Length - 1];
	//���̃t�@�C�����Ƌ�؂蕶��1�������𔲂����t�H���_�p�X���擾
	String^ FolderPath = FilePath->Substring(0, (FilePath->Length - (FileSinglePath->Length+1)));
	//\\�Ƃ�����؂蕶���Ńp�X�𕪊����A�z�񉻂���
	array<String^>^ checkFolderPathArray = FolderPath->Split('\\');
	//�t�H���_�����擾
	String^ FolderSinglePath = checkFolderPathArray[checkFolderPathArray->Length - 1];
	//���݊J���Ă���t�H���_�ƊJ�����Ƃ��Ă���t�@�C���̃t�H���_����v���Ă��邩�ǂ����𔻒�
	bool matchFlg = false;

	for each (String^ var in temp)
	{
		//��v����t�H���_������΁A�N�����t�H���_�ɒT���Ă���t�H���_������̂Ńt���O��true�ɂ���
		if (var == FolderPath) {
			matchFlg = true;
		}
	}

	//�t�H���_�����w�肵�āA�A�N�e�B�u�E�B���h�E������
	//IntPtr^ ip = windowCapture1::Form1::FindWindowByName("testFolder");
	//windowCapture1::Form1::SetForeGroundByIntPtr((IntPtr)ip);

	//�t�@�C���p�X���w�肵�āA���̃t�@�C�����܂܂��t�H���_���őO�ʂɊJ��
	//���ɊJ���Ă���ꍇ�͍ė��p���āA�őO�ʂɈړ�����
	windowCapture1::Form1::setForeWindowSelectFile(FilePath);


	//try {
	//	�w�肵���t�@�C���p�X��I��������ԂŃt�H���_���J���A�őO�ʂɂ����Ԃɂ���
	//	�������A���ɊJ���Ă���ꍇ�͓�d�ɊJ���Ă��܂�
	//	Process::Start("explorer.exe", "/select, ""C:\\Users\\chach\\Desktop\\testFolder\\test.txt");
	//	//�G�N�X�v���[�����N������܂őҋ@
	//	System::Threading::Thread::Sleep(1500);
	//}
	//catch (Exception^ e) {
	//	MessageBox::Show(e->ToString());
	//}
	//sleep�ł͂Ȃ�delay�őҋ@�����ɒ��킵�����A��肭������
	////Task^t = windowCapture1::Form1::AsyncFunction();
	////t->Wait();
	//���݂̃A�N�e�B�u�E�B���h�E���擾���āA�摜�Ƃ��ĕۑ�����
	//String^ Path = windowCapture1::Form1::GetGazouPath();
	return System::Void();
}






