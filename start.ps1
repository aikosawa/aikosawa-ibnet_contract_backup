# �h���b�O & �h���b�v�Ŏ󂯕t�����p�X���_�C�A���O�ŕ\������
Add-Type -Assembly System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show($args[0], "�t�@�C���̃p�X")

# Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# ���z���̗L����
venv\Scripts\Activate.ps1


# ���ϐ��̐ݒ�
# $env:CONFIG_FILE_PATH = "�ݒ���.xlsx���i�[����Ă���t�@�C���̃t���p�X���L�q"
# $env:TEMPLATE_FOLDER_PATH = "\\fileserver02\IBNet���\����J�p\��IBN �d�v��\�\�����ސ���\�����Ǘ�\���V�X�e���e���v���[�g��\templates"
# $env:OUTPUT_FOLDER_PATH = "�o�͐�̃t�H���_�̃t���p�X���L�q"


# �h���b�O & �h���b�v�Ŏ󂯕t�����t�@�C����main.py�ɓn���ċN��
py main.py $args[0]
# py main.py 'c:\Users\Gou\Downloads\�l���̓f�[�^ - 20230130 .xlsx'