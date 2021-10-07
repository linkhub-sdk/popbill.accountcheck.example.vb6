VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "팝빌 예금주조회 SDK Example"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14805
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   14805
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame5 
      Caption         =   "예금주조회"
      Height          =   1335
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   4935
      Begin VB.TextBox txtAccountNumber 
         Height          =   270
         Left            =   1200
         TabIndex        =   33
         Text            =   "94324511758"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtBankCode 
         Height          =   270
         Left            =   1200
         TabIndex        =   32
         Text            =   "0004"
         Top             =   320
         Width           =   1935
      End
      Begin VB.CommandButton btnCheckAccountInfo 
         Caption         =   "예금주조회"
         Height          =   855
         Left            =   3360
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "계좌번호 : "
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   31
         Top             =   765
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "기관코드 : "
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "담당자 관련"
      Height          =   1815
      Left            =   10560
      TabIndex        =   10
      Top             =   1200
      Width           =   1935
      Begin VB.CommandButton btnRegistContact 
         Caption         =   "담당자 추가"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton btnListContact 
         Caption         =   "담당자 목록 조회"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton btnUpdateContact 
         Caption         =   "담당자 정보 수정"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "포인트 관련"
      Height          =   1815
      Left            =   2040
      TabIndex        =   7
      Top             =   1200
      Width           =   1935
      Begin VB.CommandButton btnGetUnitCost 
         Caption         =   "조회단가 확인"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton btnGetChargeInfo 
         Caption         =   "과금정보 확인"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "회원정보"
      Height          =   1815
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
      Begin VB.CommandButton btnCheckID 
         Caption         =   "ID 중복 확인"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton btnCheckIsMember 
         Caption         =   "가입여부 확인"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton btnJoinMember 
         Caption         =   "회원가입"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
      End
   End
   Begin VB.TextBox txtUserCorpNum 
      Height          =   270
      Left            =   2160
      TabIndex        =   2
      Text            =   "1234567890"
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txtUserID 
      Height          =   270
      Left            =   5640
      TabIndex        =   1
      Text            =   "testkorea"
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton btnUnitCost 
      Caption         =   "조회단가 확인"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "팝빌 URL 관련"
      Height          =   1815
      Left            =   8640
      TabIndex        =   9
      Top             =   1200
      Width           =   1815
      Begin VB.CommandButton btnGetAccessURL 
         Caption         =   "팝빌 로그인 URL"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "팝빌기본 API"
      Height          =   2295
      Left            =   120
      TabIndex        =   17
      Top             =   960
      Width           =   14535
      Begin VB.Frame Frame7 
         Caption         =   "회사정보 관련"
         Height          =   1815
         Left            =   12480
         TabIndex        =   24
         Top             =   240
         Width           =   1815
         Begin VB.CommandButton btnGetCorpInfo 
            Caption         =   "회사정보 조회"
            Height          =   375
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "회사정보 수정"
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   1575
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "연동과금 포인트"
         Height          =   1815
         Left            =   4080
         TabIndex        =   21
         Top             =   240
         Width           =   2055
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "잔여포인트 확인"
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton btnGetChargeURL 
            Caption         =   "포인트 충전 URL"
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   720
            Width           =   1815
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "파트너과금 포인트"
         Height          =   1815
         Left            =   6240
         TabIndex        =   18
         Top             =   240
         Width           =   2175
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "파트너포인트 확인"
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1935
         End
         Begin VB.CommandButton btnGetPartnerURL_CHRG 
            Caption         =   "포인트 충전 URL"
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   1935
         End
      End
   End
   Begin VB.Label Label1 
      Caption         =   "팝빌회원 사업자번호 : "
      Height          =   225
      Left            =   240
      TabIndex        =   28
      Top             =   495
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "팝빌회원 아이디 : "
      Height          =   225
      Left            =   4080
      TabIndex        =   27
      Top             =   495
      Width           =   1455
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' 팝빌 예금주조회 API VB 6.0 SDK Example
'
'
' - 업데이트 일자 : 2021-10-07
' - 연동 기술지원 연락처 : 1600-9854 / 070-4504-2991
' - 연동 기술지원 이메일 : code@linkhub.co.kr
'
' <테스트 연동개발 준비사항>
' 1) 25, 28번 라인에 선언된 링크아이디(LinkID)와 비밀키(SecretKey)를
'    링크허브 가입시 메일로 발급받은 인증정보를 참조하여 변경합니다.
' 2) 팝빌 개발용 사이트(test.popbill.com)에 연동회원으로 가입합니다.
'=========================================================================

Option Explicit

'=========================================================================
' - 인증정보(링크아이디, 비밀키)는 파트너의 연동회원을 식별하는
'   인증에 사용되는 정보로 유출되지 않도록 주의하시기 바랍니다.
' - 상업용 전환이후에도 인증정보(링크아이디, 비밀키)는 변경되지 않습니다.
'=========================================================================

'링크아이디
Private Const linkID = "TESTER"

'비밀키. 유출에 주의하시기 바랍니다.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

'예금주조회 서비스 객체 생성
Private AccountCheckService As New PBAccountCheckService

'=========================================================================
' 1건의 예금주성명을 조회합니다.
' - https://docs.popbill.com/accountcheck/vb/api#CheckAccountInfo
'=========================================================================
Private Sub btnCheckAccountInfo_Click()
    Dim AccountInfo As PBAccountCheckInfo
    Dim tmp As String
    
    Set AccountInfo = AccountCheckService.CheckAccountInfo(txtUserCorpNum.Text, txtBankCode.Text, txtAccountNumber.Text)
    
    If AccountInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "응답메시지 : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "bankCode (기관코드) : " + AccountInfo.bankCode + vbCrLf
    tmp = tmp + "accountNumber (계좌번호) : " + AccountInfo.accountNumber + vbCrLf
    tmp = tmp + "accountName (예금주 성명) : " + AccountInfo.accountName + vbCrLf
    tmp = tmp + "checkDate (확인일시) : " + AccountInfo.checkDate + vbCrLf
    tmp = tmp + "resultCode (응답코드) : " + AccountInfo.resultCode + vbCrLf
    tmp = tmp + "resultMessage (응답메시지) : " + AccountInfo.resultMessage
    
    MsgBox tmp, , "예금주조회"
End Sub

'=========================================================================
' 사용하고자 하는 아이디의 중복여부를 확인합니다.
' - https://docs.popbill.com/accountcheck/vb/api#CheckID
'=========================================================================
Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = AccountCheckService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "응답메시지 : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 사업자번호를 조회하여 연동회원 가입여부를 확인합니다.
' - LinkID는 인증정보로 설정되어 있는 링크아이디 값입니다.
' - https://docs.popbill.com/accountcheck/vb/api#CheckIsMember
'=========================================================================
Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = AccountCheckService.CheckIsMember(txtUserCorpNum.Text, linkID)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "응답메시지 : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 팝빌 사이트에 로그인 상태로 접근할 수 있는 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/accountcheck/vb/api#GetAccessURL
'=========================================================================
Private Sub btnGetAccessURL_Click()
    Dim url As String
    
    url = AccountCheckService.GetAccessURL(txtUserCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "응답메시지 : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 연동회원의 잔여포인트를 확인합니다.
' - 과금방식이 파트너과금인 경우 파트너 잔여포인트(GetPartnerBalance API)를 통해 확인하시기 바랍니다.
' - https://docs.popbill.com/accountcheck/vb/api#GetBalance
'=========================================================================
Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = AccountCheckService.GetBalance(txtUserCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("응답코드 : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "응답메시지 : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "연동회원 잔여포인트 : " + CStr(balance)
End Sub

'=========================================================================
' 예금주조회 API 서비스 과금정보를 확인합니다.
' - https://docs.popbill.com/accountcheck/vb/api#GetChargeInfo
'=========================================================================
Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBchargeInfo
    Dim tmp As String
    
    Set ChargeInfo = AccountCheckService.GetChargeInfo(txtUserCorpNum.Text)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "응답메시지 : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "unitCost (조회단가) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (과금유형) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (과금제도) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원 포인트 충전을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/accountcheck/vb/api#GetChargeURL
'=========================================================================
Private Sub btnGetChargeURL_Click()
    Dim url As String
    
    url = AccountCheckService.GetChargeURL(txtUserCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "응답메시지 : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 연동회원의 회사정보를 확인합니다.
' - https://docs.popbill.com/accountcheck/vb/api#GetCorpInfo
'=========================================================================
Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    Dim tmp As String
    
    Set CorpInfo = AccountCheckService.GetCorpInfo(txtUserCorpNum.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "응답메시지 : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "ceoname (대표자성명) : " + CorpInfo.CEOName + vbCrLf
    tmp = tmp + "corpName (상호명) : " + CorpInfo.CorpName + vbCrLf
    tmp = tmp + "addr (주소) : " + CorpInfo.Addr + vbCrLf
    tmp = tmp + "bizType (업태) : " + CorpInfo.BizType + vbCrLf
    tmp = tmp + "bizClass (종목) : " + CorpInfo.BizClass + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 파트너의 잔여포인트를 확인합니다.
' - 과금방식이 연동과금인 경우 연동회원 잔여포인트(GetBalance API)를 이용하시기 바랍니다.
' - https://docs.popbill.com/accountcheck/vb/api#GetPartnerBalance
'=========================================================================
Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = AccountCheckService.GetPartnerBalance(txtUserCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("응답코드 : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "응답메시지 : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "파트너 잔여포인트 : " + CStr(balance)
End Sub

'=========================================================================
' 파트너 포인트 충전을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/accountcheck/vb/api#GetPartnerURL
'=========================================================================
Private Sub btnGetPartnerURL_CHRG_Click()
    Dim url As String
    
    url = AccountCheckService.GetPartnerURL(txtUserCorpNum.Text, "CHRG")
       
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "응답메시지 : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    
End Sub

'=========================================================================
' 예금주 성명 조회시 과금되는 포인트 단가를 확인합니다.
' - https://docs.popbill.com/accountcheck/vb/api#GetUnitCost
'=========================================================================
Private Sub btnGetUnitCost_Click()
    Dim unitCost As Double
    
    unitCost = AccountCheckService.GetUnitCost(txtUserCorpNum.Text)
    
    If unitCost < 0 Then
        MsgBox ("응답코드 : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "응답메시지 : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "조회단가 : " + CStr(unitCost)
End Sub

'=========================================================================
' 사용자를 연동회원으로 가입처리합니다.
' - https://docs.popbill.com/accountcheck/vb/api#JoinMember
'=========================================================================
Private Sub btnJoinMember_Click()

    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '아이디, 6자이상 50자 미만
    joinData.ID = "userid"
    
    '비밀번호, 6자이상 20자 미만
    joinData.PWD = "pwd_must_be_long_enough"
    
    '파트너링크 아이디
    joinData.linkID = linkID
    
    '사업자번호, '-'제외, 10자리
    joinData.CorpNum = "1234567890"
    
    '대표자성명, 최대 100자
    joinData.CEOName = "대표자성명"
    
    '상호명, 최대 200자
    joinData.CorpName = "회원상호"
    
    '사업장 주소, 최대 300자
    joinData.Addr = "주소"
    
    '업태, 최대 100자
    joinData.BizType = "업태"
    
    '종목, 최대 100자
    joinData.BizClass = "종목"

    '담당자 성명, 최대 100자
    joinData.contactName = "담당자성명"
    
    '담당자 이메일, 최대 100자
    joinData.ContactEmail = "test@test.com"
    
    '담당자 연락처, 최대 20자
    joinData.ContactTEL = "02-999-9999"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.ContactHP = "010-1234-5678"
    
    '담당자 팩스번호, 최대 20자
    joinData.ContactFAX = "02-999-9998"
    
    Set Response = AccountCheckService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "응답메시지 : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원 사업자번호에 등록된 담당자(팝빌 로그인 계정) 목록을 확인합니다.
' - https://docs.popbill.com/accountcheck/vb/api#ListContact
'=========================================================================
Private Sub btnListContact_Click()
    Dim resultList As Collection
    Dim tmp As String
    Dim info As PBContactInfo
    
    Set resultList = AccountCheckService.ListContact(txtUserCorpNum.Text)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "응답메시지 : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(아이디) | personName(성명) | email(이메일) | hp(휴대폰번호) |  fax(팩스번호) | tel(연락처) | " _
         + "regDT(등록일시) | searchAllAllowYN(회사조회 권한여부) | mgrYN(관리자 여부) | state(상태) " + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.ID + " | " + info.personName + " | " + info.email + " | " + info.hp + " | " + info.fax _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchAllAllowYN) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원 사업자번호에 담당자(팝빌 로그인 계정)를 추가합니다.
' - https://docs.popbill.com/accountcheck/vb/api#RegistContact
'=========================================================================
Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디, 6자 이상 50자 미만
    joinData.ID = "testkorea"
    
    '비밀번호, 6자 이상 20자 미만
    joinData.PWD = "test@test.com"
    
    '담당자명, 최대 100자
    joinData.personName = "담당자명"
    
    '담당자 연락처, 최대 20자
    joinData.tel = "070-1234-1234"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.hp = "010-1234-1234"
    
    '담당자 팩스번,최대 20자
    joinData.fax = "070-1234-1234"
    
    '담당자 메일주소, 최대 100자
    joinData.email = "test@test.com"
    
    '회사조회 권한여부, True-회사조회 / False-개인조회
    joinData.searchAllAllowYN = True
    
    '관리자 여부, True-관리자 / False-사용자
    joinData.mgrYN = False
    
    Set Response = AccountCheckService.RegistContact(txtUserCorpNum.Text, joinData)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "응답메시지 : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원 사업자번호에 등록된 담당자(팝빌 로그인 계정) 정보를 수정합니다.
' - https://docs.popbill.com/accountcheck/vb/api#UpdateContact
'=========================================================================
Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디
    joinData.ID = txtUserID.Text
    
    '담당자 성명, 최대 100자
    joinData.personName = "담당자명_수정"
    
    '담당자 연락처, 최대 20자
    joinData.tel = "070-1234-1234"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.hp = "010-1234-1234"
        
    '담당자 팩스번호, 최대 20자
    joinData.fax = "070-1234-1234"
    
    '담당자 이메일, 최대 100자
    joinData.email = "test@test.com"

    '회사조회 권한여부, True-회사조회 / False-개인조회
    joinData.searchAllAllowYN = True
    
    '관리자 여부, True-관리자 / False-사용자
    joinData.mgrYN = False
                
    Set Response = AccountCheckService.UpdateContact(txtUserCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "응답메시지 : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원의 회사정보를 수정합니다.
' - https://docs.popbill.com/accountcheck/vb/api#UpdateCorpInfo
'=========================================================================
Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '대표자명, 최대 100자
    CorpInfo.CEOName = "대표자"
    
    '상호, 최대 200자
    CorpInfo.CorpName = "상호"
    
    '주소, 최대 300자
    CorpInfo.Addr = "서울특별시"
    
    '업태, 최대 100자
    CorpInfo.BizType = "업태"
    
    '종목, 최대 100자
    CorpInfo.BizClass = "종목"
    
    Set Response = AccountCheckService.UpdateCorpInfo(txtUserCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "응답메시지 : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

Private Sub Form_Load()

    '모듈 초기화
    AccountCheckService.Initialize linkID, SecretKey
    
    '연동환경 설정값 True(개발용), False(상업용)
    AccountCheckService.IsTest = True
    
    '인증토큰 IP제한기능 사용여부, True(권장)
    AccountCheckService.IPRestrictOnOff = True
    
    ' 팝빌 API 서비스 고정 IP 사용여부, True-사용, False-미사용, 기본값(False)
    AccountCheckService.UseStaticIP = False
    
    ' 로컬시스템 시간 사용여부 True-사용, Fasle-미사용, 기본값(False)
    AccountCheckService.UseLocalTimeYN = False
    
End Sub

