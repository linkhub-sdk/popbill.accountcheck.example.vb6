VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "�˺� ��������ȸ SDK Example"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14805
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   14805
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.Frame Frame5 
      Caption         =   "��������ȸ"
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
         Caption         =   "��������ȸ"
         Height          =   855
         Left            =   3360
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "���¹�ȣ : "
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   31
         Top             =   765
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "����ڵ� : "
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "����� ����"
      Height          =   1815
      Left            =   10560
      TabIndex        =   10
      Top             =   1200
      Width           =   1935
      Begin VB.CommandButton btnRegistContact 
         Caption         =   "����� �߰�"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton btnListContact 
         Caption         =   "����� ��� ��ȸ"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton btnUpdateContact 
         Caption         =   "����� ���� ����"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "����Ʈ ����"
      Height          =   1815
      Left            =   2040
      TabIndex        =   7
      Top             =   1200
      Width           =   1935
      Begin VB.CommandButton btnGetUnitCost 
         Caption         =   "��ȸ�ܰ� Ȯ��"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton btnGetChargeInfo 
         Caption         =   "�������� Ȯ��"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ȸ������"
      Height          =   1815
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
      Begin VB.CommandButton btnCheckID 
         Caption         =   "ID �ߺ� Ȯ��"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton btnCheckIsMember 
         Caption         =   "���Կ��� Ȯ��"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton btnJoinMember 
         Caption         =   "ȸ������"
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
      Caption         =   "��ȸ�ܰ� Ȯ��"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "�˺� URL ����"
      Height          =   1815
      Left            =   8640
      TabIndex        =   9
      Top             =   1200
      Width           =   1815
      Begin VB.CommandButton btnGetAccessURL 
         Caption         =   "�˺� �α��� URL"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "�˺��⺻ API"
      Height          =   2295
      Left            =   120
      TabIndex        =   17
      Top             =   960
      Width           =   14535
      Begin VB.Frame Frame7 
         Caption         =   "ȸ������ ����"
         Height          =   1815
         Left            =   12480
         TabIndex        =   24
         Top             =   240
         Width           =   1815
         Begin VB.CommandButton btnGetCorpInfo 
            Caption         =   "ȸ������ ��ȸ"
            Height          =   375
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "ȸ������ ����"
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   1575
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "�������� ����Ʈ"
         Height          =   1815
         Left            =   4080
         TabIndex        =   21
         Top             =   240
         Width           =   2055
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "�ܿ�����Ʈ Ȯ��"
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton btnGetChargeURL 
            Caption         =   "����Ʈ ���� URL"
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   720
            Width           =   1815
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "��Ʈ�ʰ��� ����Ʈ"
         Height          =   1815
         Left            =   6240
         TabIndex        =   18
         Top             =   240
         Width           =   2175
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "��Ʈ������Ʈ Ȯ��"
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1935
         End
         Begin VB.CommandButton btnGetPartnerURL_CHRG 
            Caption         =   "����Ʈ ���� URL"
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   1935
         End
      End
   End
   Begin VB.Label Label1 
      Caption         =   "�˺�ȸ�� ����ڹ�ȣ : "
      Height          =   225
      Left            =   240
      TabIndex        =   28
      Top             =   495
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "�˺�ȸ�� ���̵� : "
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
' �˺� ��������ȸ API VB 6.0 SDK Example
'
'
' - ������Ʈ ���� : 2021-10-07
' - ���� ������� ����ó : 1600-9854 / 070-4504-2991
' - ���� ������� �̸��� : code@linkhub.co.kr
'
' <�׽�Ʈ �������� �غ����>
' 1) 25, 28�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
'    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
' 2) �˺� ���߿� ����Ʈ(test.popbill.com)�� ����ȸ������ �����մϴ�.
'=========================================================================

Option Explicit

'=========================================================================
' - ��������(��ũ���̵�, ���Ű)�� ��Ʈ���� ����ȸ���� �ĺ��ϴ�
'   ������ ���Ǵ� ������ ������� �ʵ��� �����Ͻñ� �ٶ��ϴ�.
' - ����� ��ȯ���Ŀ��� ��������(��ũ���̵�, ���Ű)�� ������� �ʽ��ϴ�.
'=========================================================================

'��ũ���̵�
Private Const linkID = "TESTER"

'���Ű. ���⿡ �����Ͻñ� �ٶ��ϴ�.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

'��������ȸ ���� ��ü ����
Private AccountCheckService As New PBAccountCheckService

'=========================================================================
' 1���� �����ּ����� ��ȸ�մϴ�.
' - https://docs.popbill.com/accountcheck/vb/api#CheckAccountInfo
'=========================================================================
Private Sub btnCheckAccountInfo_Click()
    Dim AccountInfo As PBAccountCheckInfo
    Dim tmp As String
    
    Set AccountInfo = AccountCheckService.CheckAccountInfo(txtUserCorpNum.Text, txtBankCode.Text, txtAccountNumber.Text)
    
    If AccountInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "����޽��� : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "bankCode (����ڵ�) : " + AccountInfo.bankCode + vbCrLf
    tmp = tmp + "accountNumber (���¹�ȣ) : " + AccountInfo.accountNumber + vbCrLf
    tmp = tmp + "accountName (������ ����) : " + AccountInfo.accountName + vbCrLf
    tmp = tmp + "checkDate (Ȯ���Ͻ�) : " + AccountInfo.checkDate + vbCrLf
    tmp = tmp + "resultCode (�����ڵ�) : " + AccountInfo.resultCode + vbCrLf
    tmp = tmp + "resultMessage (����޽���) : " + AccountInfo.resultMessage
    
    MsgBox tmp, , "��������ȸ"
End Sub

'=========================================================================
' ����ϰ��� �ϴ� ���̵��� �ߺ����θ� Ȯ���մϴ�.
' - https://docs.popbill.com/accountcheck/vb/api#CheckID
'=========================================================================
Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = AccountCheckService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "����޽��� : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ڹ�ȣ�� ��ȸ�Ͽ� ����ȸ�� ���Կ��θ� Ȯ���մϴ�.
' - LinkID�� ���������� �����Ǿ� �ִ� ��ũ���̵� ���Դϴ�.
' - https://docs.popbill.com/accountcheck/vb/api#CheckIsMember
'=========================================================================
Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = AccountCheckService.CheckIsMember(txtUserCorpNum.Text, linkID)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "����޽��� : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �˺� ����Ʈ�� �α��� ���·� ������ �� �ִ� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/accountcheck/vb/api#GetAccessURL
'=========================================================================
Private Sub btnGetAccessURL_Click()
    Dim url As String
    
    url = AccountCheckService.GetAccessURL(txtUserCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "����޽��� : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' ����ȸ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ��Ʈ�ʰ����� ��� ��Ʈ�� �ܿ�����Ʈ(GetPartnerBalance API)�� ���� Ȯ���Ͻñ� �ٶ��ϴ�.
' - https://docs.popbill.com/accountcheck/vb/api#GetBalance
'=========================================================================
Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = AccountCheckService.GetBalance(txtUserCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "����޽��� : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "����ȸ�� �ܿ�����Ʈ : " + CStr(balance)
End Sub

'=========================================================================
' ��������ȸ API ���� ���������� Ȯ���մϴ�.
' - https://docs.popbill.com/accountcheck/vb/api#GetChargeInfo
'=========================================================================
Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBchargeInfo
    Dim tmp As String
    
    Set ChargeInfo = AccountCheckService.GetChargeInfo(txtUserCorpNum.Text)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "����޽��� : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "unitCost (��ȸ�ܰ�) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (��������) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (��������) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ ������ ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/accountcheck/vb/api#GetChargeURL
'=========================================================================
Private Sub btnGetChargeURL_Click()
    Dim url As String
    
    url = AccountCheckService.GetChargeURL(txtUserCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "����޽��� : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� Ȯ���մϴ�.
' - https://docs.popbill.com/accountcheck/vb/api#GetCorpInfo
'=========================================================================
Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    Dim tmp As String
    
    Set CorpInfo = AccountCheckService.GetCorpInfo(txtUserCorpNum.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "����޽��� : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "ceoname (��ǥ�ڼ���) : " + CorpInfo.CEOName + vbCrLf
    tmp = tmp + "corpName (��ȣ��) : " + CorpInfo.CorpName + vbCrLf
    tmp = tmp + "addr (�ּ�) : " + CorpInfo.Addr + vbCrLf
    tmp = tmp + "bizType (����) : " + CorpInfo.BizType + vbCrLf
    tmp = tmp + "bizClass (����) : " + CorpInfo.BizClass + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' ��Ʈ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ���������� ��� ����ȸ�� �ܿ�����Ʈ(GetBalance API)�� �̿��Ͻñ� �ٶ��ϴ�.
' - https://docs.popbill.com/accountcheck/vb/api#GetPartnerBalance
'=========================================================================
Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = AccountCheckService.GetPartnerBalance(txtUserCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "����޽��� : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "��Ʈ�� �ܿ�����Ʈ : " + CStr(balance)
End Sub

'=========================================================================
' ��Ʈ�� ����Ʈ ������ ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/accountcheck/vb/api#GetPartnerURL
'=========================================================================
Private Sub btnGetPartnerURL_CHRG_Click()
    Dim url As String
    
    url = AccountCheckService.GetPartnerURL(txtUserCorpNum.Text, "CHRG")
       
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "����޽��� : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    
End Sub

'=========================================================================
' ������ ���� ��ȸ�� ���ݵǴ� ����Ʈ �ܰ��� Ȯ���մϴ�.
' - https://docs.popbill.com/accountcheck/vb/api#GetUnitCost
'=========================================================================
Private Sub btnGetUnitCost_Click()
    Dim unitCost As Double
    
    unitCost = AccountCheckService.GetUnitCost(txtUserCorpNum.Text)
    
    If unitCost < 0 Then
        MsgBox ("�����ڵ� : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "����޽��� : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "��ȸ�ܰ� : " + CStr(unitCost)
End Sub

'=========================================================================
' ����ڸ� ����ȸ������ ����ó���մϴ�.
' - https://docs.popbill.com/accountcheck/vb/api#JoinMember
'=========================================================================
Private Sub btnJoinMember_Click()

    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '���̵�, 6���̻� 50�� �̸�
    joinData.ID = "userid"
    
    '��й�ȣ, 6���̻� 20�� �̸�
    joinData.PWD = "pwd_must_be_long_enough"
    
    '��Ʈ�ʸ�ũ ���̵�
    joinData.linkID = linkID
    
    '����ڹ�ȣ, '-'����, 10�ڸ�
    joinData.CorpNum = "1234567890"
    
    '��ǥ�ڼ���, �ִ� 100��
    joinData.CEOName = "��ǥ�ڼ���"
    
    '��ȣ��, �ִ� 200��
    joinData.CorpName = "ȸ����ȣ"
    
    '����� �ּ�, �ִ� 300��
    joinData.Addr = "�ּ�"
    
    '����, �ִ� 100��
    joinData.BizType = "����"
    
    '����, �ִ� 100��
    joinData.BizClass = "����"

    '����� ����, �ִ� 100��
    joinData.contactName = "����ڼ���"
    
    '����� �̸���, �ִ� 100��
    joinData.ContactEmail = "test@test.com"
    
    '����� ����ó, �ִ� 20��
    joinData.ContactTEL = "02-999-9999"
    
    '����� �޴�����ȣ, �ִ� 20��
    joinData.ContactHP = "010-1234-5678"
    
    '����� �ѽ���ȣ, �ִ� 20��
    joinData.ContactFAX = "02-999-9998"
    
    Set Response = AccountCheckService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "����޽��� : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� ��ϵ� �����(�˺� �α��� ����) ����� Ȯ���մϴ�.
' - https://docs.popbill.com/accountcheck/vb/api#ListContact
'=========================================================================
Private Sub btnListContact_Click()
    Dim resultList As Collection
    Dim tmp As String
    Dim info As PBContactInfo
    
    Set resultList = AccountCheckService.ListContact(txtUserCorpNum.Text)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "����޽��� : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(���̵�) | personName(����) | email(�̸���) | hp(�޴�����ȣ) |  fax(�ѽ���ȣ) | tel(����ó) | " _
         + "regDT(����Ͻ�) | searchAllAllowYN(ȸ����ȸ ���ѿ���) | mgrYN(������ ����) | state(����) " + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.ID + " | " + info.personName + " | " + info.email + " | " + info.hp + " | " + info.fax _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchAllAllowYN) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� �����(�˺� �α��� ����)�� �߰��մϴ�.
' - https://docs.popbill.com/accountcheck/vb/api#RegistContact
'=========================================================================
Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�, 6�� �̻� 50�� �̸�
    joinData.ID = "testkorea"
    
    '��й�ȣ, 6�� �̻� 20�� �̸�
    joinData.PWD = "test@test.com"
    
    '����ڸ�, �ִ� 100��
    joinData.personName = "����ڸ�"
    
    '����� ����ó, �ִ� 20��
    joinData.tel = "070-1234-1234"
    
    '����� �޴�����ȣ, �ִ� 20��
    joinData.hp = "010-1234-1234"
    
    '����� �ѽ���,�ִ� 20��
    joinData.fax = "070-1234-1234"
    
    '����� �����ּ�, �ִ� 100��
    joinData.email = "test@test.com"
    
    'ȸ����ȸ ���ѿ���, True-ȸ����ȸ / False-������ȸ
    joinData.searchAllAllowYN = True
    
    '������ ����, True-������ / False-�����
    joinData.mgrYN = False
    
    Set Response = AccountCheckService.RegistContact(txtUserCorpNum.Text, joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "����޽��� : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� ��ϵ� �����(�˺� �α��� ����) ������ �����մϴ�.
' - https://docs.popbill.com/accountcheck/vb/api#UpdateContact
'=========================================================================
Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�
    joinData.ID = txtUserID.Text
    
    '����� ����, �ִ� 100��
    joinData.personName = "����ڸ�_����"
    
    '����� ����ó, �ִ� 20��
    joinData.tel = "070-1234-1234"
    
    '����� �޴�����ȣ, �ִ� 20��
    joinData.hp = "010-1234-1234"
        
    '����� �ѽ���ȣ, �ִ� 20��
    joinData.fax = "070-1234-1234"
    
    '����� �̸���, �ִ� 100��
    joinData.email = "test@test.com"

    'ȸ����ȸ ���ѿ���, True-ȸ����ȸ / False-������ȸ
    joinData.searchAllAllowYN = True
    
    '������ ����, True-������ / False-�����
    joinData.mgrYN = False
                
    Set Response = AccountCheckService.UpdateContact(txtUserCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "����޽��� : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� �����մϴ�.
' - https://docs.popbill.com/accountcheck/vb/api#UpdateCorpInfo
'=========================================================================
Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '��ǥ�ڸ�, �ִ� 100��
    CorpInfo.CEOName = "��ǥ��"
    
    '��ȣ, �ִ� 200��
    CorpInfo.CorpName = "��ȣ"
    
    '�ּ�, �ִ� 300��
    CorpInfo.Addr = "����Ư����"
    
    '����, �ִ� 100��
    CorpInfo.BizType = "����"
    
    '����, �ִ� 100��
    CorpInfo.BizClass = "����"
    
    Set Response = AccountCheckService.UpdateCorpInfo(txtUserCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(AccountCheckService.LastErrCode) + vbCrLf + "����޽��� : " + AccountCheckService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

Private Sub Form_Load()

    '��� �ʱ�ȭ
    AccountCheckService.Initialize linkID, SecretKey
    
    '����ȯ�� ������ True(���߿�), False(�����)
    AccountCheckService.IsTest = True
    
    '������ū IP���ѱ�� ��뿩��, True(����)
    AccountCheckService.IPRestrictOnOff = True
    
    ' �˺� API ���� ���� IP ��뿩��, True-���, False-�̻��, �⺻��(False)
    AccountCheckService.UseStaticIP = False
    
    ' ���ýý��� �ð� ��뿩�� True-���, Fasle-�̻��, �⺻��(False)
    AccountCheckService.UseLocalTimeYN = False
    
End Sub

