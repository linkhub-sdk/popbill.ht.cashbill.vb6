VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBHTCashbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ntsconfirmNum        As String       '����û���ι�ȣ
Public tradeDT              As String       '�ŷ��Ͻ�
Public tradeUsage           As String       '�ŷ�����
Public tradeType            As String       '���ݿ����� ����
Public supplyCost           As String       '���ް���
Public tax                  As String       '����
Public serviceFee           As String       '�����
Public totalAmount          As String       '�ŷ��ݾ�

Public franchiseCorpNum     As String       '������ ����ڹ�ȣ
Public franchiseCorpName    As String       '������ ��ȣ
Public franchiseCorpType    As String       '������ ���������

Public identityNum          As String       '�ŷ�ó �ĺ���ȣ
Public identityNumType      As String       '�ĺ���ȣ ����
Public customerName         As String       '������
Public cardOwnerName        As String       'ī������ڸ�
Public deductionType        As String       '��������

Public Sub fromDictionary(dic As Variant)
    
    ntsconfirmNum = dic.Item("ntsconfirmNum")
    tradeDT = dic.Item("tradeDT")
    tradeUsage = dic.Item("tradeUsage")
    tradeType = dic.Item("tradeType")
    supplyCost = dic.Item("supplyCost")
    tax = dic.Item("tax")
    serviceFee = dic.Item("serviceFee")
    totalAmount = dic.Item("totalAmount")
    
    franchiseCorpNum = dic.Item("franchiseCorpNum")
    franchiseCorpName = dic.Item("franchiseCorpName")
    franchiseCorpType = dic.Item("franchiseCorpType")
    
    identityNum = dic.Item("identityNum")
    identityNumType = dic.Item("identityNumType")
    customerName = dic.Item("customerName")
    cardOwnerName = dic.Item("cardOwnerName")
    deductionType = dic.Item("deductionType")
    
End Sub

