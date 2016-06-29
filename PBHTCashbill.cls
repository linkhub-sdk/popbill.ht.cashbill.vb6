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

Public ntsconfirmNum        As String       '국세청승인번호
Public tradeDT              As String       '거래일시
Public tradeUsage           As String       '거래유형
Public tradeType            As String       '현금영수증 형태
Public supplyCost           As String       '공급가액
Public tax                  As String       '세액
Public serviceFee           As String       '봉사료
Public totalAmount          As String       '거래금액

Public franchiseCorpNum     As String       '발행자 사업자번호
Public franchiseCorpName    As String       '발행자 상호
Public franchiseCorpType    As String       '발행자 사업자유형

Public identityNum          As String       '거래처 식별번호
Public identityNumType      As String       '식별번호 유형
Public customerName         As String       '고객명
Public cardOwnerName        As String       '카드소유자명
Public deductionType        As String       '공제유형

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


