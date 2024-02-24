Attribute VB_Name = "mAntiBot"
Option Explicit

Private Const FILE_PATCH_WORKING As String = "\VALIDATION\WORKING.ini"

Private Const FILE_PATCH_EMAILS  As String = "\VALIDATION\MAILS.ini"

Private Const MAX_MAILS          As Byte = 200

Private Const MAX_MAILS_FOR_DAY  As Byte = 10


Public Enum TypeWorking

    eDS_AccountNew = 1
    eDS_AccountRecover = 2
    eDS_AccountPasswd = 3
    
    eNewCharPVP = 4 ' No se usa en Desterium Code
    eDS_Mercader_New = 5
    eDS_Mercader_Offer = 6
    eDS_Mercader_New_Confirm = 7
    eDS_Mercader_NewOffer_Confirm = 8
    
End Enum
