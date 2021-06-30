Module AppGlobals
    Public strURL As String
    Public strTimiosURL As String
    Public strVisiRecordingURL As String
    Public strCommonwealthURL As String
    Public strUserId As String
    Public strPassword As String
    Public strListenPath As String
    Public strProcessedPath As String
    Public strExceptionPath As String
    Public intPollingIntervalMins As Integer
    Public strTempDir As String
    Public strLogDir As String
    Public strMailFromAddress As String
    Public strMailFromUser As String
    Public strMailFromPassword As String
    Public strMailToUser1 As String
    Public strMailToUser2 As String
    Public oServiceLog As ServiceLog
    Public intProcessing As Integer = 0
    Public intUploadTimeout As Integer
    Public intVisiRecordingTimeout As Integer
    Public intWFGError As Integer = 1
    Public intVSIError As Integer = 1
    Public intVLRError As Integer = 1
    Public intVisiRecordingError As Integer = 1
    Public intTimiosError As Integer = 1
    Public intReswareError As Integer = 1
    Public intCommonwealthError As Integer = 1
    Public intSilkReswareError As Integer = 1
    Public intHowardHannaError As Integer = 1
    Public intAllError As Integer = 1
    Public blnUploadToVSI As Boolean = False
    Public blnUploadToVLR As Boolean = False
    Public blnUploadToVisiRecording As Boolean = False
    Public blnUploadToWFG As Boolean = False
    Public blnUploadToResware As Boolean = False
    Public blnUploadToHowardHanna As Boolean = False
    Public blnLoggedInToResware As Boolean = False
    Public blnUploadToCommonwealth As Boolean = False
    Public blnUploadToSilkResware As Boolean = False
    Public blnUploadToTimios As Boolean = False
    Public objScanningRecord As ScanningRecord
End Module
