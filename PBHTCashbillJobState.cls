VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBHTCashbillJobState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public jobID            As String   '�۾����̵�
Public jobState         As Integer  '�۾�����
Public queryType        As String   '����(����)��꼭 ����
Public queryDateType    As String   '�������� ����
Public queryStDate      As String   '��������
Public queryEnDate      As String   '��������
Public errorCode        As Long     '�����ڵ�, �����ΰ�� 1��ȯ
Public errorReason      As String   '�����޽���
Public jobStartDT       As String   'Ȩ�ý� ���� �����Ͻ�
Public jobEndDT         As String   'Ȩ�ý� ���� �����Ͻ�
Public collectCount     As Integer  '������� �Ǽ�
Public regDT            As String   '���� ��û�Ͻ�

Public Sub fromDictionary(dic As Variant)

    jobID = dic.Item("jobID")
    jobState = dic.Item("jobState")
    queryType = dic.Item("queryType")
    queryDateType = dic.Item("queryDateType")
    queryStDate = dic.Item("queryStDate")
    queryEnDate = dic.Item("queryEnDate")
    errorCode = dic.Item("errorCode")
    errorReason = dic.Item("errorReason")
    jobStartDT = dic.Item("jobStartDT")
    jobEndDT = dic.Item("jobEndDT")
    collectCount = dic.Item("collectCount")
    regDT = dic.Item("regDT")
End Sub


