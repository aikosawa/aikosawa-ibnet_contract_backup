# ���ł�venv������ꍇ�ɂ͍폜
if (Test-Path .\venv) {
    Remove-Item -Path .\venv -Recurse
}

# �O���[�o���̃p�b�P�[�W��j�󂵂Ȃ����߂ɉ��z���𒼉���venv�f�B���N�g���ɍ쐬
py -m venv venv 

# ���z���̗L����
venv\Scripts\Activate.ps1

# �ˑ����C�u�����̃C���X�g�[��
pip install -r requirements.txt