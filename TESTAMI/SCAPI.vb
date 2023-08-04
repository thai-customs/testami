Imports System.Runtime.InteropServices
Imports System.Text

Public Class SCAPI
    '// ---- AID ---- //

    Public Const CM_AID As String = "434D"
    Public Const ADM_AID As String = "A000000084060002"
    Public Const MOI_AID As String = "A000000054480001"

    '// ---- FUNCTION DECLARATION ---- //

    Public Declare Function LicenseManager Lib "scapi_ope.dll" (ByRef status As Integer) As Short

    Public Declare Function SetDebugOption Lib "scapi_ope.dll" (ByVal debug_level As Integer) As Short

    Public Declare Function ListReader Lib "scapi_ope.dll" (ByVal List_Reader As String, ByRef status As Integer) As Short

    Public Declare Function OpenReader Lib "scapi_ope.dll" (ByVal reader_name As String, ByRef status As Integer) As Short

    Public Declare Function CloseReader Lib "scapi_ope.dll" () As Short

    Public Declare Function GetCardStatus Lib "scapi_ope.dll" (ByVal atr As String, ByRef atrLen As Integer, ByVal timeOut As Integer, ByRef cardType As Integer, ByRef status As Integer) As Short

    Public Declare Function SelectApplet Lib "scapi_ope.dll" (ByRef aid As Byte, ByVal aid_size As Integer, ByRef status As Integer) As Short

    Public Declare Function GetCardInfo Lib "scapi_ope.dll" (ByVal card_sn As String, ByVal chip As String, ByVal os As String, ByVal pre_perso As String, ByVal perso As String, ByRef status As Integer) As Short

    '// --- FUNCTION DECLARATION --- //

    Public Declare Function GetInfoADM Lib "scapi_ope.dll" (ByVal version As String, ByRef status As Integer, ByRef authorize As Integer, ByVal laser_number As String, ByRef status As Integer) As Short

    Public Declare Function GetInfoACM Lib "scapi_ope.dll" (ByVal version As String, ByRef auth_status As Integer, ByRef status As Integer) As Short

    Public Declare Function GetInfoTemplate Lib "scapi_ope.dll" (ByVal version As String, ByRef status As Integer, ByVal share_data As String, ByRef status As Integer) As Short

    Public Declare Function ReadData Lib "scapi_ope.dll" (ByVal block_id As Integer, ByVal offset As Integer, ByVal data_size As Integer, ByVal data_buf As String, ByRef status As Integer) As Short

    Public Declare Function GetMatchStatus Lib "scapi_ope.dll" (ByVal req_type As Integer, ByVal req_mode As Integer, ByVal in_buf As String, ByVal in_size As Integer, ByVal out_buf As String, ByRef out_size As Integer, ByRef match_stt As Integer, ByRef status As Integer) As Short

    Public Declare Function SwitchAuthorize Lib "scapi_ope.dll" (ByRef status As Integer) As Short

    Public Declare Function DerivePIN Lib "scapi_ope.dll" (ByVal pin As String, ByVal pin_len As Integer, ByVal dpin As String, ByRef dpin_len As Integer) As Short

    Public Declare Function EnvelopeGMS2 Lib "scapi_ope.dll" (ByVal atr As String, ByVal atr_len As Integer, ByVal cid As String, ByVal cid_len As Integer, ByVal key_id As Integer, ByVal cryptogram As String, ByVal cryptogram_len As Integer, ByVal request As String, ByRef request_len As Integer) As Short

    Public Declare Function VerifyPIN Lib "scapi_ope.dll" (ByVal pin_id As Integer, ByVal share_data As Integer, ByRef try_remain As Integer, ByRef status As Integer) As Short

    Public Declare Function ChangePIN Lib "scapi_ope.dll" (ByVal pin_id As Integer, ByRef try_remain As Integer, ByRef status As Integer) As Short

    Public Declare Function OverwritePIN1 Lib "scapi_ope.dll" (ByRef status As Integer) As Short

    Public Declare Function VerifyPIN_OBJ Lib "scapi_ope.dll" (ByVal pin_id As Integer, ByVal pin_oid As String, ByVal share_data As Integer, ByRef try_remain As Integer, ByRef status As Integer) As Short

    Public Declare Function ChangePIN_OBJ Lib "scapi_ope.dll" (ByVal pin_id As Integer, ByVal cpin_oid As String, ByVal npin1_oid As String, ByVal npin2_oid As String, ByRef try_remain As Integer, ByRef status As Integer) As Short

    Public Declare Function OverwritePIN1_OBJ Lib "scapi_ope.dll" (ByVal pin_oid As String, ByRef status As Integer) As Short

    Public Declare Function EnvelopeGMSx Lib "scapi_ope.dll" (ByVal key_id As Integer, ByVal cryptogram As String, ByVal cryptogram_len As Integer, ByVal request As String, ByRef request_len As Integer, ByRef status As Integer) As Short

    '---------------------------------------------
    '   START  SCAPI_SASM   : SASM function  
    '   BEGIN            : 2008-05-20
    '   UPDATED        : 2008-05-28
    '---------------------------------------------

    '- INTERNAL AUTH KEY
    Public Const SAS_INT_AUTH_RMACKEY_ADMIN = 0
    Public Const SAS_INT_AUTH_FPKEY_ADMIN = 1
    Public Const SAS_INT_AUTH_FPKEY_MOI = 2
    Public Const SAS_INT_AUTH_FPKEY_HTH = 3
    Public Const SAS_INT_AUTH_FPKEY_MOI2 = 4
    Public Const SAS_INT_AUTH_FPKEY_MOI3 = 5
    Public Const SAS_INT_AUTH_FPKEY_MOI4 = 6
    Public Const SAS_INT_AUTH_FPKEY_WVOT = 7
    Public Const SAS_INT_AUTH_FPKEY_MOD = 8
    Public Const SAS_INT_AUTH_FPKEY_MOA = 9
    Public Const SAS_INT_AUTH_FPKEY_NHSO = 10

    '- SESSION BUFFER SIZE

    Public Const MAC_LENGTH = 4

    'Public Declare Function EnvelopeGMS Lib "scapi_ope.dll" (
    'ByVal atr As String,
    'ByVal atr_len As Long,
    'ByVal cid As String,
    'ByVal cid_len As Long,
    'ByVal key_id As Long,
    'ByVal cryptogram As String,
    'ByVal cryptogram_len As Long,
    'ByVal request As String,
    'ByRef request_len As Long) As Integer

    Public Declare Function GenerateMAC Lib "scapi_ope.dll" (
    ByVal cid As String,
    ByVal buf As String,
    ByVal buf_len As Integer,
    ByVal mac As String,
    ByRef mac_len As Integer,
    ByRef status As Integer) As Integer

    Public Declare Function VerifyMAC Lib "scapi_ope.dll" (
    ByVal cid As String,
    ByVal buf As String,
    ByVal buf_len As Integer,
    ByVal mac As String,
    ByRef mac_len As Integer,
    ByRef status As Integer) As Integer

    Public Declare Function GetRMAC Lib "scapi_ope.dll" (
    ByVal ibuf As String,
    ByVal ibuf_len As Integer,
    ByVal obuf As String,
    ByRef obuf_len As Integer,
    ByRef status As Integer) As Integer

    Public Declare Function SetIgnoreCard Lib "scapi_ope.dll" () As Integer

    ' ------------------------
    ' END  SCAPI_SASM   : SASM function
    ' ------------------------
    Public Class SCAPI_STATUS

        Public Const APC_FAILED As Long = -9L
        Public Const APDU_ERROR As Long = -1L
        Public Const IRIS_API_ERROR As Long = -4L
        Public Const PB_API_ERROR As Long = -6L
        Public Const PB_E_FINALIZE_FAILED As Long = 54L
        Public Const PB_E_INTERNAL_MAX As Long = 21L
        Public Const PB_E_INTERNAL_MIN As Long = 1L
        Public Const PB_E_STORE_BIO_HEADER_FAILED As Long = 52L
        Public Const PB_E_STORE_REF_DATA_FAILED As Long = 53L
        Public Const PB_VALIDATE_SCORE_INSUFFICIENT As Long = 51L
        Public Const SAS_FAILED As Long = -5L
        Public Const SAS_STATUS_ALGO_NOT_AVAILABLE As Long = 4100L
        Public Const SAS_STATUS_AUTHEN_FAILED As Long = 4112L
        Public Const SAS_STATUS_INVALID_MESSAGE As Long = 4098L
        Public Const SAS_STATUS_KEY_NOT_AVAILABLE As Long = 4106L
        Public Const SAS_STATUS_REQUEST_FAILED As Long = 4097L
        Public Const SAS_STATUS_VERIFY_FAILED As Long = 4110L
        Public Const SCAPI_BIOCARDAPI_FAILED As Long = -3L
        Public Const SCAPI_BRC100T_FAILED As Long = -4L
        Public Const SCAPI_FAILED As Long = -1L
        Public Const SCAPI_SCARD_FAILED As Long = -2L
        Public Const SCAPI_STATUS_APPLICATION_INVALIDATED As Long = 1102L
        Public Const SCAPI_STATUS_APPLICATION_NOT_EXIST As Long = 1101L
        Public Const SCAPI_STATUS_AUTHENTICATION_FAILED As Long = 1011L
        Public Const SCAPI_STATUS_CARD_LOCKED As Long = 1501L
        Public Const SCAPI_STATUS_CID_NOT_FOUND As Long = 3011L
        Public Const SCAPI_STATUS_COMMUNICATION_ERROR As Long = 3002L
        Public Const SCAPI_STATUS_CONDITION_NOT_SATISFIED As Long = 1024L
        Public Const SCAPI_STATUS_ENCODING_ERROR As Long = 3012L
        Public Const SCAPI_STATUS_ENCODING_UNKNOWN_ERROR As Long = 3099L
        Public Const SCAPI_STATUS_FUNCTION_NOT_SUPPORT As Long = 1201L
        Public Const SCAPI_STATUS_IDCARD_INTERNAL_ERROR As Long = 1998L
        Public Const SCAPI_STATUS_IDCARD_UNKNOWN_ERROR As Long = 1999L
        Public Const SCAPI_STATUS_INCORRECT_PIN As Long = 1001L
        Public Const SCAPI_STATUS_INPUT_INCORRECT As Long = 1202L
        Public Const SCAPI_STATUS_KEY_CURRENTLY_BLOCKED As Long = 1012L
        Public Const SCAPI_STATUS_NEW_PIN_NOT_MATCH As Long = 1007L
        Public Const SCAPI_STATUS_NO_CARD_PRESENT As Long = 2002L
        Public Const SCAPI_STATUS_NO_LICENSE_MANAGER As Long = 9001L
        Public Const SCAPI_STATUS_NO_PERMISSION As Long = 1022L
        Public Const SCAPI_STATUS_NO_PERMIT_FROM_CARD_HOLDER As Long = 1003L
        Public Const SCAPI_STATUS_NOT_FP_AUTHORIZE As Long = 1004L
        Public Const SCAPI_STATUS_NOT_PIN_AUTHORIZE As Long = 1005L
        Public Const SCAPI_STATUS_PIN_CURRENTLY_BLOCKED As Long = 1002L
        Public Const SCAPI_STATUS_PINBOX_OBJ_ERROR As Long = 1006L
        Public Const SCAPI_STATUS_READER_NOT_OPEN_YET As Long = 2003L
        Public Const SCAPI_STATUS_REFERENCE_DATA_INVALID As Long = 1211L
        Public Const SCAPI_STATUS_REFERENCE_DATA_NOT_FOUND As Long = 1212L
        Public Const SCAPI_STATUS_SAME_OR_UNKNOWN_OR_INAPPROPIATE_STATUS As Long = 1103L
        Public Const SCAPI_STATUS_SECURITY_STATUS_NOT_SATISFIED As Long = 1023L
        Public Const SCAPI_STATUS_SYSTEM_CANCELLED As Long = 3001L
        Public Const SCAPI_STATUS_UNKNOWN_CARD_TYPE As Long = 2004L
        Public Const SCAPI_STATUS_UNKNOWN_READER As Long = 2001L
        Public Const SCAPI_STATUS_WRONG_OPTION As Long = 1203L
        Public Const SCAPI_SUCCESS As Long = 0L
        'Public Const SUCCESS As Long = 0L
        Public Const SW_ALGO_NOT_AVAILABLE As Long = 28435L
        Public Const SW_APP_NOT_AVAILABLE As Long = 28434L
        Public Const SW_CARD_NOT_SUPPORTED As Long = 28448L
        Public Const SW_GEN_PKI_KEY_FAILED As Long = 28428L
        Public Const SW_INSTALL_NOT_AVAILABLE As Long = 28433L
        Public Const SW_KEY_INCORRECT As Long = 28427L
        Public Const SW_KEY_NOT_AVAILABLE As Long = 28426L
        Public Const SW_LOADFILE_NOT_AVAILABLE As Long = 28432L
        Public Const SW_NOT_AVAILABLE As Long = 28424L
        Public Const SW_NOT_ENABLED As Long = 28569L
        Public Const SW_NOT_MATCH As Long = 28431L
        Public Const SW_NOT_SUPPORTED As Long = 28425L
        Public Const SW_PIN_NOT_AVAILABLE As Long = 28442L
        Public Const SW_RANDOM_NOT_AVAILABLE As Long = 28423L
        Public Const SW_SEED_LOCKED As Long = 28421L
        Public Const SW_SEED_NOT_LOCKED As Long = 28422L
        Public Const SW_SERVICE_NOT_AUTHENTICATED As Long = 27380L
        Public Const SW_SERVICE_NOT_AVAILABLE As Long = 27377L
        Public Const SW_SERVICE_NOT_ENABLED As Long = 27379L
        Public Const SW_SERVICE_NOT_SUPPORTED As Long = 27378L
        Public Const SW_SIGN_FAILED As Long = 28429L
        Public Const SW_VERIFY_FAILED As Long = 28430L
        Public Const SW_WRONG_ALGORITHM As Long = 28418L
        Public Const SW_WRONG_RANDOM As Long = 28419L
        Public Const SW_WRONG_SAM_TYPE As Long = 28420L


        Public Function GetStatus(ByVal status4 As Long) As String
            Dim retOut As String = ""
            Select Case status4
                Case APC_FAILED
                    retOut = "APC FAILED"
                Case PB_API_ERROR
                    retOut = "PB API ERROR"
                Case PB_E_FINALIZE_FAILED
                    retOut = "PB E FINALIZE FAILED"
                Case PB_E_INTERNAL_MAX
                    retOut = "PB E INTERNAL MAX"
                Case PB_E_INTERNAL_MIN
                    retOut = "PB E INTERNAL MIN"
                Case PB_E_STORE_BIO_HEADER_FAILED
                    retOut = "PB E STORE BIO HEADER FAILED"
                Case PB_E_STORE_REF_DATA_FAILED
                    retOut = "PB E STORE REF DATA FAILED"
                Case PB_VALIDATE_SCORE_INSUFFICIENT
                    retOut = "PB VALIDATE SCORE INSUFFICIENT"
                Case SAS_FAILED
                    retOut = "SAS FAILED"
                Case SAS_STATUS_ALGO_NOT_AVAILABLE
                    retOut = "SAS STATUS ALGO NOT AVAILABLE"
                Case SAS_STATUS_AUTHEN_FAILED
                    retOut = "SAS STATUS AUTHEN FAILED"
                Case SAS_STATUS_INVALID_MESSAGE
                    retOut = "SAS STATUS INVALID MESSAGE"
                Case SAS_STATUS_KEY_NOT_AVAILABLE
                    retOut = "SAS STATUS KEY NOT AVAILABLE"
                Case SAS_STATUS_REQUEST_FAILED
                    retOut = "SAS STATUS REQUEST FAILED"
                Case SAS_STATUS_VERIFY_FAILED
                    retOut = "SAS STATUS VERIFY FAILED"
                Case SCAPI_BIOCARDAPI_FAILED
                    retOut = "SCAPI BIOCARDAPI FAILED"
                Case SCAPI_BRC100T_FAILED
                    retOut = "SCAPI BRC100T FAILED"
                Case SCAPI_FAILED
                    retOut = "SCAPI FAILED"
                Case SCAPI_SCARD_FAILED
                    retOut = "SCAPI SCARD FAILED"
                Case SCAPI_STATUS_APPLICATION_INVALIDATED
                    retOut = "SCAPI STATUS APPLICATION INVALIDATED"
                Case SCAPI_STATUS_APPLICATION_NOT_EXIST
                    retOut = "SCAPI STATUS APPLICATION NOT EXIST"
                Case SCAPI_STATUS_AUTHENTICATION_FAILED
                    retOut = "SCAPI STATUS AUTHENTICATION FAILED"
                Case SCAPI_STATUS_CARD_LOCKED
                    retOut = "SCAPI STATUS CARD LOCKED"
                Case SCAPI_STATUS_CID_NOT_FOUND
                    retOut = "SCAPI STATUS CID NOT FOUND"
                Case SCAPI_STATUS_COMMUNICATION_ERROR
                    retOut = "SCAPI STATUS COMMUNICATION ERROR"
                Case SCAPI_STATUS_CONDITION_NOT_SATISFIED
                    retOut = "SCAPI STATUS CONDITION NOT SATISFIED"
                Case SCAPI_STATUS_ENCODING_ERROR
                    retOut = "SCAPI STATUS ENCODING ERROR"
                Case SCAPI_STATUS_ENCODING_UNKNOWN_ERROR
                    retOut = "SCAPI STATUS ENCODING UNKNOWN ERROR"
                Case SCAPI_STATUS_FUNCTION_NOT_SUPPORT
                    retOut = "SCAPI STATUS FUNCTION NOT SUPPORT"
                Case SCAPI_STATUS_IDCARD_INTERNAL_ERROR
                    retOut = "SCAPI STATUS IDCARD INTERNAL ERROR"
                Case SCAPI_STATUS_IDCARD_UNKNOWN_ERROR
                    retOut = "SCAPI STATUS IDCARD UNKNOWN ERROR"
                Case SCAPI_STATUS_INCORRECT_PIN
                    retOut = "SCAPI STATUS INCORRECT PIN"
                Case SCAPI_STATUS_INPUT_INCORRECT
                    retOut = "SCAPI STATUS INPUT INCORRECT"
                Case SCAPI_STATUS_KEY_CURRENTLY_BLOCKED
                    retOut = "SCAPI STATUS KEY CURRENTLY BLOCKED"
                Case SCAPI_STATUS_NEW_PIN_NOT_MATCH
                    retOut = "SCAPI STATUS NEW PIN NOT MATCH"
                Case SCAPI_STATUS_NO_CARD_PRESENT
                    retOut = "SCAPI STATUS NO CARD PRESENT"
                Case SCAPI_STATUS_NO_LICENSE_MANAGER
                    retOut = "SCAPI STATUS NO LICENSE MANAGER"
                Case SCAPI_STATUS_NO_PERMISSION
                    retOut = "SCAPI STATUS NO PERMISSION"
                Case SCAPI_STATUS_NO_PERMIT_FROM_CARD_HOLDER
                    retOut = "SCAPI STATUS NO PERMIT FROM CARD HOLDER"
                Case SCAPI_STATUS_NOT_FP_AUTHORIZE
                    retOut = "SCAPI STATUS NOT FP AUTHORIZE"
                Case SCAPI_STATUS_NOT_PIN_AUTHORIZE
                    retOut = "SCAPI STATUS NOT PIN AUTHORIZE"
                Case SCAPI_STATUS_PIN_CURRENTLY_BLOCKED
                    retOut = "SCAPI STATUS PIN CURRENTLY BLOCKED"
                Case SCAPI_STATUS_PINBOX_OBJ_ERROR
                    retOut = "SCAPI STATUS PINBOX OBJ ERROR"
                Case SCAPI_STATUS_READER_NOT_OPEN_YET
                    retOut = "SCAPI STATUS READER NOT OPEN YET"
                Case SCAPI_STATUS_REFERENCE_DATA_INVALID
                    retOut = "SCAPI STATUS REFERENCE DATA INVALID"
                Case SCAPI_STATUS_REFERENCE_DATA_NOT_FOUND
                    retOut = "SCAPI STATUS REFERENCE DATA NOT FOUND"
                Case SCAPI_STATUS_SAME_OR_UNKNOWN_OR_INAPPROPIATE_STATUS
                    retOut = "SCAPI STATUS SAME OR UNKNOWN OR INAPPROPIATE STATUS"
                Case SCAPI_STATUS_SECURITY_STATUS_NOT_SATISFIED
                    retOut = "SCAPI STATUS SECURITY STATUS NOT SATISFIED"
                Case SCAPI_STATUS_SYSTEM_CANCELLED
                    retOut = "SCAPI STATUS SYSTEM CANCELLED"
                Case SCAPI_STATUS_UNKNOWN_CARD_TYPE
                    retOut = "SCAPI STATUS UNKNOWN CARD TYPE"
                Case SCAPI_STATUS_UNKNOWN_READER
                    retOut = "SCAPI STATUS UNKNOWN READER"
                Case SCAPI_STATUS_WRONG_OPTION
                    retOut = "SCAPI STATUS WRONG OPTION"
                Case SCAPI_SUCCESS
                    retOut = "SCAPI SUCCESS"
                Case SW_ALGO_NOT_AVAILABLE
                    retOut = "SW ALGO NOT AVAILABLE"
                Case SW_APP_NOT_AVAILABLE
                    retOut = "SW APP NOT AVAILABLE"
                Case SW_CARD_NOT_SUPPORTED
                    retOut = "SW CARD NOT SUPPORTED"
                Case SW_GEN_PKI_KEY_FAILED
                    retOut = "SW GEN PKI KEY FAILED"
                Case SW_INSTALL_NOT_AVAILABLE
                    retOut = "SW INSTALL NOT AVAILABLE"
                Case SW_KEY_INCORRECT
                    retOut = "SW KEY INCORRECT"
                Case SW_KEY_NOT_AVAILABLE
                    retOut = "SW KEY NOT AVAILABLE"
                Case SW_LOADFILE_NOT_AVAILABLE
                    retOut = "SW LOADFILE NOT AVAILABLE"
                Case SW_NOT_AVAILABLE
                    retOut = "SW NOT AVAILABLE"
                Case SW_NOT_ENABLED
                    retOut = "SW NOT ENABLED"
                Case SW_NOT_MATCH
                    retOut = "SW NOT MATCH"
                Case SW_NOT_SUPPORTED
                    retOut = "SW NOT SUPPORTED"
                Case SW_PIN_NOT_AVAILABLE
                    retOut = "SW PIN NOT AVAILABLE"
                Case SW_RANDOM_NOT_AVAILABLE
                    retOut = "SW RANDOM NOT AVAILABLE"
                Case SW_SEED_LOCKED
                    retOut = "SW SEED LOCKED"
                Case SW_SEED_NOT_LOCKED
                    retOut = "SW SEED NOT LOCKED"
                Case SW_SERVICE_NOT_AUTHENTICATED
                    retOut = "SW SERVICE NOT AUTHENTICATED"
                Case SW_SERVICE_NOT_AVAILABLE
                    retOut = "SW SERVICE NOT AVAILABLE"
                Case SW_SERVICE_NOT_ENABLED
                    retOut = "SW SERVICE NOT ENABLED"
                Case SW_SERVICE_NOT_SUPPORTED
                    retOut = "SW SERVICE NOT SUPPORTED"
                Case SW_SIGN_FAILED
                    retOut = "SW SIGN FAILED"
                Case SW_VERIFY_FAILED
                    retOut = "SW VERIFY FAILED"
                Case SW_WRONG_ALGORITHM
                    retOut = "SW WRONG ALGORITHM"
                Case SW_WRONG_RANDOM
                    retOut = "SW WRONG RANDOM"
                Case SW_WRONG_SAM_TYPE
                    retOut = "SW WRONG SAM TYPE"
                Case Else
                    retOut = "Unknow Status" '& " [" + status4.ToString() & "]"
                    'retOut = AMI.GetStatus3(status4)
            End Select

            Return retOut
        End Function

    End Class


End Class

Public Class AMI


    Public Declare Function AMI_SESSION Lib "AMI32.DLL" Alias "AMI_SESSION" (
            ByVal Session_Id As String,
            ByRef size As Integer) As Integer

    'Public Declare Function AMI_REQUEST Lib "AMI32.DLL" Alias "AMI_REQUEST" (
    '        ByVal HOSTNAME As String,
    '        ByVal send_data As String,
    '        ByRef send_size As Integer,
    '        ByVal reply_data() As Byte,
    '        ByRef reply_max As Integer,
    '        ByRef reply_size As Integer,
    '        ByRef TimeLimit As Integer,
    '        ByRef status As Integer) As Integer

    Public Declare Function AMI_REQUEST Lib "AMI32.DLL" Alias "AMI_REQUEST" (
            ByVal HOSTNAME As String,
            ByVal send_data As String,
            ByRef send_size As Integer,
            ByVal reply_data As IntPtr,
            ByRef reply_max As Integer,
            ByRef reply_size As Integer,
            ByRef TimeLimit As Integer,
            ByRef status As Integer) As Integer

    'Public Declare Function AMI_REQUEST Lib "AMI32.DLL" Alias "AMI_REQUEST" (
    '        ByVal HOSTNAME() As Byte,
    '        ByVal send_data() As Byte,
    '        ByRef send_size As Integer,
    '        ByVal reply_data() As Byte,
    '        ByRef reply_max As Integer,
    '        ByRef reply_size As Integer,
    '        ByRef TimeLimit As Integer,
    '        ByRef status As Integer) As Integer


    Public Structure Recive9080
        ReadOnly ReActioCode4 As String
        ReadOnly ReturnStatus5 As String
        ReadOnly XKey32 As String
        ReadOnly XKeyBit As String
        Public Sub New(ByRef amiRCV() As Byte)
            Dim st As String = Encoding.Default.GetString(amiRCV)
            ReActioCode4 = st.Substring(0, 4)
            ReturnStatus5 = st.Substring(4, 5)
            XKey32 = st.Substring(9, 32)
            Dim bb(st.Length - 9) As Byte
            Buffer.BlockCopy(amiRCV, 9, bb, 0, st.Length - 9)
            XKeyBit = BitConverter.ToString(bb).ToString().Replace("-", "")
        End Sub
    End Structure

    Public Structure Recive9081
        ReadOnly ReActioCode4 As String
        ReadOnly ReturnStatus5 As String
        ReadOnly TKey32 As String
        ReadOnly TKeyBit As String
        Public Sub New(ByRef amiRCV() As Byte)
            Dim st As String = Encoding.Default.GetString(amiRCV)
            ReActioCode4 = st.Substring(0, 4)
            ReturnStatus5 = st.Substring(4, 5)
            TKey32 = st.Substring(9)
            Dim bb(st.Length - 9) As Byte
            Buffer.BlockCopy(amiRCV, 9, bb, 0, st.Length - 9)
            TKeyBit = BitConverter.ToString(bb).ToString().Replace("-", "")
        End Sub
    End Structure

    Public Structure Recive5090
        ReadOnly ReActioCode4 As String
        ReadOnly ReturnStatus5 As String
        Public Sub New(ByRef amiRCV() As Byte)
            Dim st As String = Encoding.Default.GetString(amiRCV)
            ReActioCode4 = st.Substring(0, 4)
            ReturnStatus5 = st.Substring(4)
        End Sub
    End Structure

    Public Structure Recive5000
        ReadOnly ReActioCode4 As String
        ReadOnly ReturnStatus5 As String
        ReadOnly DataJSON As String
        Public Sub New(ByRef amiRCV() As Byte)
            Dim st As String = Encoding.UTF8.GetString(amiRCV)
            ReActioCode4 = st.Substring(0, 4)
            ReturnStatus5 = st.Substring(4, 5)
            If st.Length > 9 Then
                DataJSON = st.Substring(9)
            End If
        End Sub
    End Structure
    '<StructLayout(LayoutKind.Sequential)>

    Public Structure GetPOP
        <VBFixedArray(4), MarshalAs(UnmanagedType.ByValArray, SizeConst:=4)>
        Public Reqno() As Byte '0101
        <VBFixedArray(9), MarshalAs(UnmanagedType.ByValArray, SizeConst:=9)>
        Public ReqId() As Byte 'Space(9)
        <VBFixedArray(4), MarshalAs(UnmanagedType.ByValArray, SizeConst:=4)>
        Public ReqPw() As Byte 'Space(4)
        <VBFixedArray(48), MarshalAs(UnmanagedType.ByValArray, SizeConst:=48)>
        Public ReqKey() As Byte '18-65
        <VBFixedArray(1), MarshalAs(UnmanagedType.ByValArray, SizeConst:=1)>
        Public ReplyCode() As Byte '66
        <VBFixedArray(1), MarshalAs(UnmanagedType.ByValArray, SizeConst:=1)>
        Public ReqLevel() As Byte '67
        <VBFixedArray(61), MarshalAs(UnmanagedType.ByValArray, SizeConst:=61)>
        Public ActiveField() As Byte ' 68-128
        <VBFixedArray(13), MarshalAs(UnmanagedType.ByValArray, SizeConst:=13)>
        Public ReqPID() As Byte
        <VBFixedArray(16), MarshalAs(UnmanagedType.ByValArray, SizeConst:=16)>
        Public ReqCID() As Byte
        <VBFixedArray(32), MarshalAs(UnmanagedType.ByValArray, SizeConst:=32)>
        Public ReqXXX() As Byte 'binary xxx
        '**********************************
        'Acccenter As String * 1
        'Privilege As String * 18
    End Structure
    Public Structure ReplyPOP
        <VBFixedArray(4), MarshalAs(UnmanagedType.ByValArray, SizeConst:=4)>
        Public Reqno() As Byte
        <VBFixedArray(5), MarshalAs(UnmanagedType.ByValArray, SizeConst:=5)>
        Public returnCode() As Byte
        <VBFixedArray(64), MarshalAs(UnmanagedType.ByValArray, SizeConst:=64)>
        Public ActiveField() As Byte
        <VBFixedArray(13), MarshalAs(UnmanagedType.ByValArray, SizeConst:=13)>
        Public PID() As Byte
        <VBFixedArray(30), MarshalAs(UnmanagedType.ByValArray, SizeConst:=30)>
        Public Title() As Byte
        <VBFixedArray(24), MarshalAs(UnmanagedType.ByValArray, SizeConst:=24)>
        Public FNAME() As Byte
        <VBFixedArray(24), MarshalAs(UnmanagedType.ByValArray, SizeConst:=24)>
        Public LNAME() As Byte
        <VBFixedArray(4), MarshalAs(UnmanagedType.ByValArray, SizeConst:=4)>
        Public Sex() As Byte
        <VBFixedArray(8), MarshalAs(UnmanagedType.ByValArray, SizeConst:=8)>
        Public Dob() As Byte
        <VBFixedArray(30), MarshalAs(UnmanagedType.ByValArray, SizeConst:=30)>
        Public Nat() As Byte
        <VBFixedArray(24), MarshalAs(UnmanagedType.ByValArray, SizeConst:=24)>
        Public HStat() As Byte
        <VBFixedArray(24), MarshalAs(UnmanagedType.ByValArray, SizeConst:=24)>
        Public PStat() As Byte
        <VBFixedArray(8), MarshalAs(UnmanagedType.ByValArray, SizeConst:=8)>
        Public DMoveIn() As Byte
        <VBFixedArray(3), MarshalAs(UnmanagedType.ByValArray, SizeConst:=3)>
        Public age() As Byte
        <VBFixedArray(13), MarshalAs(UnmanagedType.ByValArray, SizeConst:=13)>
        Public FPID() As Byte
        <VBFixedArray(13), MarshalAs(UnmanagedType.ByValArray, SizeConst:=13)>
        Public MPID() As Byte
        <VBFixedArray(24), MarshalAs(UnmanagedType.ByValArray, SizeConst:=24)>
        Public FFName() As Byte
        <VBFixedArray(24), MarshalAs(UnmanagedType.ByValArray, SizeConst:=24)>
        Public MFName() As Byte
        <VBFixedArray(30), MarshalAs(UnmanagedType.ByValArray, SizeConst:=30)>
        Public FNat() As Byte
        <VBFixedArray(30), MarshalAs(UnmanagedType.ByValArray, SizeConst:=30)>
        Public MNat() As Byte
        <VBFixedArray(30), MarshalAs(UnmanagedType.ByValArray, SizeConst:=30)>
        Public ChangeNat() As Byte
        <VBFixedArray(8), MarshalAs(UnmanagedType.ByValArray, SizeConst:=8)>
        Public DChangeNat() As Byte
        <VBFixedArray(11), MarshalAs(UnmanagedType.ByValArray, SizeConst:=11)>
        Public Hid() As Byte
        <VBFixedArray(220), MarshalAs(UnmanagedType.ByValArray, SizeConst:=220)>
        Public HDesc() As Byte
        <VBFixedArray(2048), MarshalAs(UnmanagedType.ByValArray, SizeConst:=2048)>
        Public Reserve() As Byte
    End Structure
    Public Declare Function AMI_REQUEST Lib "AMI32.DLL" Alias "AMI_REQUEST" (
            ByVal HOSTNAME As String,
            ByRef send_data As GetPOP,
            ByRef send_size As Integer,
            ByRef reply_data As ReplyPOP,
            ByRef reply_imax As Integer,
            ByRef reply_size As Integer,
            ByRef TimeLimit As Integer,
            ByRef status As Integer) As Short

    Public Structure Get5000
        <VBFixedArray(4), MarshalAs(UnmanagedType.ByValArray, SizeConst:=4)>
        Public Code() As Byte
        <VBFixedArray(32), MarshalAs(UnmanagedType.ByValArray, SizeConst:=32)>
        Public XKey() As Byte
        <VBFixedArray(5), MarshalAs(UnmanagedType.ByValArray, SizeConst:=5)>
        Public OfficeCode() As Byte
        <VBFixedArray(2), MarshalAs(UnmanagedType.ByValArray, SizeConst:=2)>
        Public VersionCode() As Byte
        <VBFixedArray(3), MarshalAs(UnmanagedType.ByValArray, SizeConst:=3)>
        Public ServiceCode() As Byte
        <VBFixedArray(13), MarshalAs(UnmanagedType.ByValArray, SizeConst:=13)>
        Public PID() As Byte
    End Structure
    Public Structure Reply5000
        <VBFixedArray(4), MarshalAs(UnmanagedType.ByValArray, SizeConst:=4)>
        Public Reqno() As Byte
        <VBFixedArray(5), MarshalAs(UnmanagedType.ByValArray, SizeConst:=5)>
        Public returnCode() As Byte
        <VBFixedArray(32759), MarshalAs(UnmanagedType.ByValArray, SizeConst:=32759)>
        Public Data() As Byte
    End Structure
    Public Declare Function AMI_REQUEST Lib "AMI32.DLL" Alias "AMI_REQUEST" (
            ByVal HOSTNAME As String,
            ByRef send_data As Get5000,
            ByRef send_size As Integer,
            ByRef reply_data As Reply5000,
            ByRef reply_imax As Integer,
            ByRef reply_size As Integer,
            ByRef TimeLimit As Integer,
            ByRef status As Integer) As Short

    Public Class AMI_STATUS

        Public Const REQUEST_TIMEOUT As Long = 101
        Public Const CONNECT_LOST As Long = 102
        Public Const CANT_START_SERVER As Long = 103
        Public Const CONNECTION_FULL As Long = 105
        Public Const CANT_CONNECT As Long = 107
        Public Const NO_REQUEST_CONNECTION As Long = 109
        Public Const SEND_ERROR As Long = 111
        Public Const RECV_ERROR As Long = 113
        Public Const MESSAGE_TOO_LARGE As Long = 115
        Public Const UNKNOWN_MODE As Long = 201
        Public Const WRONG_MODE As Long = 203


        Public Function GetStatus3(ByVal status3 As Long) As String
            Dim retOut As String = ""
            Select Case status3
                Case REQUEST_TIMEOUT
                    retOut = "REQUEST TIMEOUT"
                Case CONNECT_LOST
                    retOut = "CONNECT LOST"
                Case CANT_START_SERVER
                    retOut = "CAN'T START SERVER"
                Case CONNECTION_FULL
                    retOut = "CONNECTION FULL"
                Case CANT_CONNECT
                    retOut = "CAN'T CONNECT"
                Case NO_REQUEST_CONNECTION
                    retOut = "NO REQUEST CONNECTION"
                Case SEND_ERROR
                    retOut = "SEND ERROR"
                Case RECV_ERROR
                    retOut = "RECV ERROR"
                Case MESSAGE_TOO_LARGE
                    retOut = "MESSAGE TOO LARGE"
                Case UNKNOWN_MODE
                    retOut = "UNKNOWN MODE"
                Case WRONG_MODE
                    retOut = "WRONG MODE"
                Case Else
                    retOut = "Unknow Status"
            End Select
            Return retOut '& " [" + status3.ToString() & "]"
        End Function

        Public Function GetReturnStatus5(ByVal status5 As Long) As String
            Dim retOut As String = ""
            If status5 > 0 And status5 < 1000 Then
                retOut = status5.ToString() + " HTTP Status จากบริการข้อมูลของหน่วยงานเชื่อมโยง (ดูความหมายของ HTTP Status Code)"
            Else
                Select Case status5
                    Case 0
                        retOut = "สำเร็จ"
                    Case 90001
                        retOut = "ไม่ได้ Login เข้าใช้งานระบบ"
                    Case 90005
                        retOut = "ไม่มีสิทธิในการทำงาน (Time to work Out)"
                    Case 90007
                        retOut = "ไม่มีสิทธิในการทำงาน (Invalid Secret Code)"
                    Case 90008
                        retOut = "ใช้สิทธิในการตรวจสอบข้อมูลครบแล้ว (Quota out of limit)"
                    Case 90009
                        retOut = "ไม่มีสิทธิในการทำงาน (Invalid Smart Card - not found in card_ctl)"
                    Case 90011
                        retOut = "ข้อมูลที่ส่งมาตรวจสอบไม่ถูกต้อง"
                    Case 90012
                        retOut = "ใช้เวลาในการให้บริการมากกว่าที่กำหนดไว้ (Service time out)"
                    Case 90013
                        retOut = "ยังไม่เปิดให้บริการหน่วยงานที่ร้องขอ"
                    Case 90014
                        retOut = "อยู่ระหว่างปรับปรุง"
                    Case 90015
                        retOut = "ไม่เปิดให้บริการ Linkage Center"
                    Case 90016
                        retOut = "Citizen not login"
                    Case 90090
                        retOut = "PIN ไม่ถูกต้อง"
                    Case 90020
                        retOut = "ไม่สามารถทำการค้นหาข้อมูลได้ (Key as Space)"
                    Case 90025
                        retOut = "ไม่สามารถเพิ่มจำนวนการใช้งานได้ (Update Account Error)"
                    Case 90026
                        retOut = "ไม่สามารถติดต่อฐานข้อมูลตรวจสอบสิทธิได้ (Connect DB Check X error)"
                    Case 90027
                        retOut = "ไม่สามารถจัดเก็บรหัสตรวจสอบข้อมูลได้ (Update X error)"
                    Case 90028
                        retOut = "ไม่สามารถตรวจสอบฐานข้อมูลบัตรของผู้ใช้งานได้ (Select iknoemp_card error)"
                    Case 90029
                        retOut = "ไม่สามารถจัดเก็บรหัสตรวจสอบข้อมูลได้ (Update X space error)"
                    Case 90040
                        retOut = "ไม่มีสิทธิในการทำงาน (card_st error)"
                    Case 90043
                        retOut = "ไม่มีสิทธิในการทำงาน (Using Code not match in emp_card)"
                    Case 90044
                        retOut = "ไม่มีสิทธิในการทำงาน (Check SAS error)"
                    Case 90045
                        retOut = "ไม่มีสิทธิในการทำงาน (SAS error - not match)"
                    Case 90046
                        retOut = "ไม่มีสิทธิในการทำงาน (Using Code sened not match in emp_card)"
                    Case 90050
                        retOut = "ไม่มีสิทธิในการทำงาน (Quota as zero)"
                    Case 90500
                        retOut = "ไม่พบข้อมูล"
                    Case 91001
                        retOut = "ไม่สามารถหาข้อมูล EDUCATION ได้"
                    Case 91002
                        retOut = "ไม่สามารถหาข้อมูล SOLDIRE ได้"
                    Case 95001
                        retOut = "ยังไม่ได้ลงทะเบียนผู้ใช้งาน"
                    Case 95002
                        retOut = "บัตรที่เข้าใช้งาน ไม่ใช่บัตรใบล่าสุด"
                    Case 95003
                        retOut = "ค่า Y ที่ส่งมาตรวจสอบไม่ถูกต้อง"
                    Case 95004
                        retOut = "ไม่มีสิทธิในการใช้งานระบบ Linkage Center"
                    Case 95011
                        retOut = "ไม่มีสิทธิในการร้องขอข้อมูลไปยัง Service ปลายทาง"
                    Case 95012
                        retOut = "ไม่พบ Service ที่ร้องขอข้อมูลในระบบ Linkage Center"
                    Case 95013
                        retOut = "ไม่สามารถร้องขอข้อมูลจาก Linkage Center ไปยัง Service ปลายทางได้"
                    Case 95014
                        retOut = "ไม่มีสิทธิในการร้องขอข้อมูลไปยัง Service ปลายทางด้วยสิทธิประชาชน"
                    Case 99301
                        retOut = "ไม่สามารถตรวจสอบภาพใบหน้าคนต่างด้าวได้"
                    Case 99304
                        retOut = "ไม่พบภาพใบหน้าคนต่างด้าวในฐานข้อมูล"
                    Case 99305
                        retOut = "ไม่สามารถอ่านไฟล์ภาพใบหน้าคนต่างด้าวได้"
                    Case 99701
                        retOut = "ไม่พบรายการการเปลี่ยนแปลงที่อยู่ในฐานข้อมูล"
                    Case 99702
                        retOut = "ไม่พบรายการการเปลี่ยนแปลงชื่อในฐานข้อมูล"
                    Case 99703
                        retOut = "ไม่พบรายการการเปลี่ยนแปลงสัญชาติในฐานข้อมูล"
                    Case 99706
                        retOut = "ไม่สามารถเพิ่มจำนวนการใช้งานได้ (Update Account Error)"
                    Case 99707
                        retOut = "ไม่พบรายการชื่อภาษาอังกฤษในฐานข้อมูล"
                    Case 99801
                        retOut = "ไม่ระบุค่าของบ้านเลขที่และรหัสจังหวัด อำเภอตำบล"
                    Case 99983
                        retOut = "ไม่สามารถส่งภาพใบหน้าได้เนื่องจากมีขนาดเกิน 20 KB"
                    Case 99989
                        retOut = "ไม่สามารถตรวจสอบรายการบัตรได้"
                    Case 99990
                        retOut = "ตรวจสอบข้อมูลแถบแม่เหล็ก เลขควบคุม 2 ไม่ถูกต้อง"
                    Case 99991
                        retOut = "ไม่พบรายการบัตรก่อนหน้า/ถัดไปในฐานข้อมูล"
                    Case 99992
                        retOut = "ไม่พบรายการบัตรในฐานข้อมูล"
                    Case 99993
                        retOut = "ไม่สามารถอ่านไฟล์ภาพใบหน้าได้"
                    Case 99994
                        retOut = "ไม่พบภาพใบหน้าในฐานข้อมูล"
                    Case 99995
                        retOut = "ไม่สามารถตรวจสอบภาพใบหน้าได้"
                    Case 99997
                        retOut = "ไม่สามารถตรวจสอบเลขควบคุมบัตรประจำตัวประชาชนได้"
                    Case 99999
                        retOut = "ไม่มีสิทธิในการทำงาน (Invalid Application ID)"
                    Case Else
                        retOut = "Unknow Status"
                End Select
            End If


            Return retOut '& " [" + status3.ToString() & "]"
        End Function
    End Class
End Class


''' <summary>
''' สำนักทะเบียน บริการข้อมูลภาพใบหน้า (ตามบัตรล่าสุด)
''' </summary>
Public Class LK_00023_01_038
    Public Property image As String
    Public Property mineType As String
End Class