Attribute VB_Name = "VBMTRun"
Option Explicit

'�����߳̾��
Public VBThreadHandle1 As Long, VBThreadHandle2 As Long
'�����߳�ID
Public VBThreadID1 As Long, VBThreadID2 As Long

'************************************ע�⣺VB6���̱߳�����SUB MAINΪ��������***************************************************
'***************************��ʾ�����Ѿ����ú��ˣ��Լ�ʹ��ʱע���ڹ��̡������ԡ�����������������ѡ��************************
Sub Main()
    If AvoidReentrant = False Then                                              '��ֹ���߳��ظ�����
        AvoidReentrant = True
        If App.PrevInstance Then                                                '��ֹ�����ظ�����
            MessageBox ByVal 0&, StrPtr("�����������л�δ��ȫ�˳�"), StrPtr("�ظ�����"), vbCritical
            Exit Sub
        Else
            InitCommonControls                                                  '��ʼ��ͨ�ÿؼ�
            GETVBHeader                                                         '��ȡVB����ͷ
            
            Form21.Show                                                         '���������������
        End If
    End If
End Sub

Public Sub Thread1()                                                            '���߳�1
    '***********************************����Ҫ����VB6�̻߳�����ʼ��*************************************************
    CreateIExprSrvObj 0&, 4&, 0&                                                'VB6���п��ʼ��
    CoInitializeEx ByVal 0&, ByVal (COINIT_MULTITHREADED Or COINIT_SPEED_OVER_MEMORY) 'COM�����ʼ��
    InitVBdll                                                                   '�յ�VB6���п��ڲ��������ֵĳ�ʼ��
    '***********************************����Ҫ����VB6�̻߳�����ʼ��*************************************************
    Dim StrTemp As String, n As Long
    Dim MyForm As Form19
    Set MyForm = New Form19                                                     '��¡FORM19������
    MyForm.Show
    '###########################################
    '#                                         #
    '#     �ǵ���û�������߳��Ҹ���û��      #
    '#           ��Ϊ������㹻����            #
    '#                                         #
    '###########################################
    CoUninitialize                                                              'ж��COM�����ʡ��Ҳ����Ӱ���ȶ��ԣ���������ɾ�����ڴ�й©��Ϊ�����ɺ�ϰ�ߣ�����д�ϣ�
End Sub
