Function BrowseForFolder(Optional sTitle As String = "�t�H���_�̑I��", Optional vRootFolder As Variant) As String
   Dim objFolder As Object
   Set objFolder = CreateObject("Shell.Application").BrowseForFolder(hWndAccessApp, sTitle, &H10, vRootFolder)
   If Not objFolder Is Nothing Then
      BrowseForFolder = objFolder.Self.Path
      Set objFolder = Nothing
   End If
End Function

' BIF_RETURNONLYFSDIRS	&H1	���t�@�C���V�X�e���̃t�H���_�[
' BIF_DONTGOBELOWDOMAIN	&H2	���l�b�g���[�N�t�H���_���܂߂Ȃ�
' BIF_STATUSTEXT	&H4	�~�X�e�[�^�X�e�L�X�g��ݒ�
' * �vCallback�iAPI SHBrowseForFolder�d�l�j
' BIF_RETURNFSANCESTORS	&H8	�����[�g�I��s��
' BIF_EDITBOX	&H10	�����̃{�b�N�X�\��
' BIF_VALIDATE	&H20	�~�I���A�C�e���̑Ó����`�F�b�N
' * �vCallback�iAPI SHBrowseForFolder�d�l�j
' BIF_NEWDIALOGSTYLE	&H40	�~�V�����X�^�C���\��
' BIF_BROWSEINCLUDEURLS	&H80	��URL��Ώۂɂł���
' BIF_UAHINT	&H100	���q���g��\���i�������ω����Ȃ��̂Ō��ʂ͊��҂ł����j
' BIF_NONEWFOLDERBUTTON	&H200	���u�V�����t�H���_�̍쐬�v��\�����Ȃ�
' BIF_NOTRANSLATETARGETS	&H400	���V���[�g�J�b�g�̃^�[�Q�b�g��PIDL��Ԃ��܂��c�ڍוs��
' BIF_BROWSEFORCOMPUTER	&H1000	���l�b�g���[�N���Ώ�
' CSIDL_NETWORK (18) ���p
' BIF_BROWSEFORPRINTER	&H2000	���v�����^�[��Ώ�
' CSIDL_PRINTERS (4) ���p
' �v�����^�[��I�����Ă��G���[�ł��i�Ӗ��Ȃ��j

' BIF_BROWSEINCLUDEFILES	&H4000	���t�@�C�����\������
' �t�@�C����I�����Ă��G���[�ł��i�Ӗ��Ȃ��j

' BIF_SHAREABLE	&H8000	�����L�\�ȃ��\�[�X��\���ł���
' �i�Ȃ����d�l���f�ڂ�ZIP,LZH�Ȃǈ��k�t�@�C���W�J������j
' �ʏ�́������W�J���邪

' ���̐ݒ肾�ƁA�������W�J����

' BIF_BROWSEFILEJUNCTIONS	&H10000	��ZIP,LZH�Ȃǈ��k���Ƀt�@�C�����\�����W�J���I�����\

' ���̑��̈��k�`�� 7z �� Rar �͕\�����ꂸ
' BIF_BROWSEINCLUDEFILES ��g�ݍ��킹�Ă��I���ł��܂���i�G���[�����j
