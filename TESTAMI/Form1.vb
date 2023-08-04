Imports TESTAMI.SCAPI
Imports System.Text
Imports System.Runtime.InteropServices
Imports Newtonsoft.Json

Public Class Form1
    Dim list_Reader As String
    Dim status As Integer = -99
    Dim returnCode As Integer = 99
    'Dim ReturnValue As Short = 99
    Dim scapi_stt As New SCAPI.SCAPI_STATUS
    Dim ami_stt As New AMI.AMI_STATUS

    Dim rc9080 As AMI.Recive9080
    Dim rc9081 As AMI.Recive9081
    Dim SID As String = Application.ProductName & Process.GetCurrentProcess.Id.ToString()
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        list_Reader = Space(1000) ' จองหน่วยความจำสำหรับเก็บชื่อ Reader ที่ได้
        returnCode = SCAPI.ListReader(list_Reader, status)
        If (returnCode = 0) Then
            list_Reader = list_Reader.Trim()
            ' ตัดคำชื่อ Reader เพิ่มเติม หากมีเครื่องอ่านมากกว่า 1 เครื่อง
            list_Reader = list_Reader.Replace(vbNullChar, "")
            While list_Reader.Length > 0
                Dim nn As String = list_Reader.Substring(0, 2)
                Dim ll As Integer = Integer.Parse(nn)

                ComboBox1.Items.Add(list_Reader.Substring(2, ll))
                ComboBox2.Items.Add(list_Reader.Substring(2, ll))

                list_Reader = list_Reader.Substring(ll + 2)
            End While
        Else
            Dim err As String = "Return Code = [" & returnCode & " ] " & Environment.NewLine & "Status Code = [ " & status & " ] " & scapi_stt.GetStatus(status) & Environment.NewLine & "ไม่พบเครื่องอ่านบัตร"
            MessageBox.Show(err, Text & " ตรวจสอบเครื่องอ่านบัตร")
            Close()
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        returnCode = 99
        status = -99
        Try
            returnCode = SCAPI.OpenReader(ComboBox1.Text, status)

            If returnCode = 0 Then
                returnCode = 99
                status = -99

                Dim atr As String = Space(100)
                Dim atr_len As Integer = 0
                Dim timeOut As Integer = 100
                Dim card_type As Integer = -999

                returnCode = SCAPI.GetCardStatus(atr, atr_len, timeOut, card_type, status)

                'TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "Get Card Status [" + returnCode.ToString() + "] " + status.ToString()
                'TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "[ART] " + atr.Trim()

                returnCode = 99
                status = -99

                Dim aid_bin(64) As Byte
                Dim aid_bin_len As Integer
                Dim util As New Utilities

                util.Str2Bin(SCAPI.MOI_AID, aid_bin, aid_bin_len)
                returnCode = SCAPI.SelectApplet(aid_bin(0), aid_bin_len, status)

                If returnCode = 0 Then
                    returnCode = 99
                    status = -99

                    Dim block_id, offset As Integer
                    Dim dataBuf As String
                    Dim data_size As Integer

                    block_id = 0
                    offset = 4
                    data_size = 13
                    dataBuf = Space(15)

                    returnCode = SCAPI.ReadData(block_id, offset, data_size, dataBuf, status)
                    TextBoxPID.Text = dataBuf

                    util.Str2Bin(SCAPI.ADM_AID, aid_bin, aid_bin_len)
                    returnCode = SCAPI.SelectApplet(aid_bin(0), aid_bin_len, status)

                    Dim cid, pre_perso, perso, chip, os As String

                    cid = Space(16)
                    pre_perso = Space(20)
                    perso = Space(20)
                    chip = Space(20)
                    os = Space(20)

                    returnCode = SCAPI.GetCardInfo(cid, chip, os, pre_perso, perso, status)
                    TextBoxCID.Text = cid
                Else

                End If
            Else
                TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "[rc=" + returnCode.ToString() + "] [st=" + status.ToString() + "] " + scapi_stt.GetStatus(status)
            End If
        Catch ex As Exception
            Dim btn As Button = CType(sender, Button)
            MessageBox.Show(ex.Message,
            btn.Text,
            MessageBoxButtons.OK,
            MessageBoxIcon.Exclamation,
            MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Dim recive_size As Integer
        Dim recive_max As Integer
        Dim timeOut As Integer = 1000
        returnCode = 99
        status = -99

        Try
            Dim send9080 As String '{0}{1}{2}{3} 9080 PID CID OfficeCode
            If TextBoxLKOfficeCode.Text.Trim().Length = 5 Then
                send9080 = String.Format("{0}{1}{2}{3}", "9080", TextBoxPID.Text.Trim(), TextBoxCID.Text.Trim(), TextBoxLKOfficeCode.Text.Trim()) '--Option :: Linkage Office Code 5 หลัก ของหน่วยงานผู้ร้องขอ
            Else
                send9080 = String.Format("{0}{1}{2}{3}", "9080", TextBoxPID.Text.Trim(), TextBoxCID.Text.Trim(), "     ") '--ใส่ช่องว่าง 5 ตัว
            End If
            recive_size = 0
            recive_max = 64 * 1024
            Dim rec_pointer As IntPtr = Marshal.AllocHGlobal(recive_max)
            returnCode = AMI.AMI_SESSION(SID, SID.Length)
            returnCode = AMI.AMI_REQUEST("", send9080, send9080.Length, rec_pointer, recive_max, recive_size, timeOut, status) '--ส่งคำสั่ง AMI

            If returnCode <> 0 Then
                TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "[rc=" + returnCode.ToString() + "] [st=" + status.ToString() + "] " + " " + ami_stt.GetStatus3(status) '--แสดงรหัสผิดพลาด กรณีพบ error
            Else
                Dim recive9080(recive_size) As Byte
                Marshal.Copy(rec_pointer, recive9080, 0, recive_size)
                rc9080 = New AMI.Recive9080(recive9080)
                If rc9080.ReturnStatus5 <> "00000" Then
                    TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + Environment.NewLine + "[rs=" + rc9080.ReturnStatus5 + "]" + ami_stt.GetReturnStatus5(Long.Parse(rc9080.ReturnStatus5))
                Else
                    TextBoxKEY.Text = rc9080.XKey32
                End If
            End If
        Catch ex As Exception
            Dim btn As Button = CType(sender, Button)
            MessageBox.Show(ex.Message,
            btn.Text,
            MessageBoxButtons.OK,
            MessageBoxIcon.Exclamation,
            MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim aid_bin(64) As Byte
        Dim aid_bin_len As Integer
        Dim util As New Utilities
        Try
            util.Str2Bin(SCAPI.ADM_AID, aid_bin, aid_bin_len)
            returnCode = SCAPI.SelectApplet(aid_bin(0), aid_bin_len, status)

            Dim adm_version, laser_number As String
            Dim authorize, adm_state As Integer

            adm_version = Space(5)
            adm_state = 0
            laser_number = Space(33)
            authorize = 0

            returnCode = 99
            status = -99
            returnCode = SCAPI.GetInfoADM(adm_version, adm_state, authorize, laser_number, status)
            If returnCode = 0 Then 'InfoADM ok
                If authorize = 0 Then 'Verify By Finger print
                    TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "Cannot PIN Function"
                Else 'authorize = 1 ,Verify By PINCODE
                    Dim try_remain As Integer
                    returnCode = 99
                    status = -99
                    returnCode = VerifyPIN(1, 0, try_remain, status)
                    If returnCode <> 0 Then
                        If status = 1001 Then
                            TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "PIN Incorrect try = " + try_remain.ToString()
                        Else
                            TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "[rc=" + returnCode.ToString() + "] [st=" + status.ToString() + "] " + scapi_stt.GetStatus(status)
                        End If
                    Else
                        Dim req_mode, req_type, match_status, random_size, cryto_size As Integer
                        Dim random, cryto As String

                        req_mode = 0
                        req_type = 1
                        random = util.Bin2Str(TextBoxKEY.Text, 32)
                        random_size = random.Length
                        cryto = Space(64)
                        cryto_size = 64
                        match_status = 0

                        returnCode = 99
                        status = -99
                        returnCode = GetMatchStatus(req_type, req_mode, random, random_size, cryto, cryto_size, match_status, status)

                        If returnCode <> 0 Then
                            TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "Get Match Status Error [st=" + status.ToString() + "] " + scapi_stt.GetStatus(status)
                        Else
                            If match_status <> 1 Then
                                TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "Match Status not match " + match_status.ToString()
                            Else
                                TextBox4.Text = cryto

                                Dim envelop As String
                                Dim envelop_size As Integer

                                envelop_size = 255
                                envelop = Space(255)

                                returnCode = 99
                                status = -99
                                returnCode = EnvelopeGMSx(SCAPI.SAS_INT_AUTH_FPKEY_ADMIN, cryto, cryto_size, envelop, envelop_size, status) 'Key SAS_INT_AUTH_FPKEY_ADMIN

                                If returnCode <> 0 Then
                                    TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "Envelop not work"
                                Else
                                    TextBox5.Text = envelop
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "[rc=" + returnCode.ToString() + "] " + "[st=" + status.ToString() + "] " + scapi_stt.GetStatus(status)
            End If

        Catch ex As Exception
            Dim btn As Button = CType(sender, Button)
            MessageBox.Show(ex.Message,
            btn.Text,
            MessageBoxButtons.OK,
            MessageBoxIcon.Exclamation,
            MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Dim reply_size As Integer
        Dim reply_max As Integer
        Dim timeOut As Integer = 1000

        Try
            '{0}{1}{2}{3}{4}:{5} --> 9081 PID CID XKey EnVelopGMSx.size : EnVelopGMSx  
            Dim send9081 As String = String.Format("{0}{1}{2}{3}{4}:{5}", "9081", TextBoxPID.Text.Trim(), TextBoxCID.Text.Trim(), TextBoxKEY.Text.Trim(), TextBox5.Text.Trim().Length.ToString(), TextBox5.Text.Trim())

            reply_size = 0
            reply_max = 64 * 1024
            Dim rec_pointer As IntPtr = Marshal.AllocHGlobal(reply_max)

            returnCode = 99
            status = -99
            returnCode = AMI.AMI_SESSION(SID, SID.Length)
            returnCode = AMI.AMI_REQUEST("", send9081, send9081.Length, rec_pointer, reply_max, reply_size, timeOut, status)

            If returnCode <> 0 Then
                TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "[rc=" + returnCode.ToString + " st=" + status.ToString() + " " + ami_stt.GetStatus3(status) + "]"
            Else
                Dim recive9081(reply_size) As Byte
                Marshal.Copy(rec_pointer, recive9081, 0, reply_size)
                rc9081 = New AMI.Recive9081(recive9081)
                TextBoxTKey.Text = rc9081.TKey32
                If rc9081.ReturnStatus5 <> "00000" Then
                    TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "[rs=" + rc9081.ReturnStatus5 + "]"
                Else
                    TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "Authorize LOGIN Success"
                End If
            End If
        Catch ex As Exception
            Dim btn As Button = CType(sender, Button)
            MessageBox.Show(ex.Message,
            btn.Text,
            MessageBoxButtons.OK,
            MessageBoxIcon.Exclamation,
            MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Try

            Dim send_data As New AMI.GetPOP
            Dim reply_data As New AMI.ReplyPOP

            Dim send_datasize As Integer
            Dim reply_datasize As Integer
            Dim reply_datamax As Integer
            Dim timeOut As Integer
            Dim status As Integer


            reply_datasize = Marshal.SizeOf(reply_data)
            reply_datamax = Marshal.SizeOf(reply_data)

            send_data.Reqno = Encoding.Default.GetBytes("0101")
            send_data.ReqId = Encoding.Default.GetBytes(New String(" ", 9))
            send_data.ReqPw = Encoding.Default.GetBytes(New String(" ", 4))

            Dim target As String
            target = TextBoxPIDex.Text.Trim()

            target = target & New String(" ", 48 - target.Length)
            send_data.ReqKey = Encoding.Default.GetBytes(target)

            send_data.ReplyCode = Encoding.Default.GetBytes("0")
            send_data.ReqLevel = Encoding.Default.GetBytes("1")
            Dim af_the_star As String = New String("1", 61)
            send_data.ActiveField = Encoding.Default.GetBytes(af_the_star)

            send_data.ReqPID = Encoding.Default.GetBytes(TextBoxPID.Text.Trim())
            send_data.ReqCID = Encoding.Default.GetBytes(TextBoxCID.Text.Trim())
            send_data.ReqXXX = Encoding.Default.GetBytes(TextBoxKEY.Text.Trim())

            send_datasize = Marshal.SizeOf(send_data)
            timeOut = 50

            returnCode = 99
            status = -99
            AMI.AMI_SESSION(SID, SID.Length)
            returnCode = AMI.AMI_REQUEST("", send_data, send_datasize, reply_data, reply_datamax, reply_datasize, timeOut, status)

            If returnCode = 0 Then
                TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + Environment.NewLine
                TextBoxOutput.Text = TextBoxOutput.Text + "PID " + Encoding.Default.GetString(reply_data.PID).Trim()
                TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine
                TextBoxOutput.Text = TextBoxOutput.Text + "ชื่อ-สกุล " + Encoding.Default.GetString(reply_data.FNAME).Trim() + " " + Encoding.Default.GetString(reply_data.LNAME).Trim()
                TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine
                TextBoxOutput.Text = TextBoxOutput.Text + "ที่อยู่ " + Encoding.Default.GetString(reply_data.HDesc).Trim().Replace("#", " ").Replace("  ", " ")
                TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine
            Else
                TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine
                TextBoxOutput.Text = TextBoxOutput.Text + "[rc=" + returnCode.ToString() + "] [st=" + status.ToString() + "] " + ami_stt.GetStatus3(status)
            End If
        Catch ex As Exception
            Dim btn As Button = CType(sender, Button)
            MessageBox.Show(ex.Message,
            btn.Text,
            MessageBoxButtons.OK,
            MessageBoxIcon.Exclamation,
            MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Sub ButtonSend5000_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSend5000.Click
        TextBoxOutput.Text = ""
        Application.DoEvents()
        Dim reply_datasize As Integer
        Dim reply_datamax As Integer
        Dim timeOut As Integer
        Dim status As Integer
        Try

            'Dim ss As Integer = AMI.AMI_SESSION(SID, SID.Length)
            'If ss > 0 Then

            returnCode = 99
                status = -99

                reply_datasize = 0
                reply_datamax = 64 * 1024

                Dim send5000 As String = String.Format("{0}{1}{2}{3}{4}{5}",
                                                                       "5000",
                                                                       TextBoxTKey.Text.Trim(), ' TKey
                                                                       TextBoxOfficeCode5.Text.Trim(), ' Office 5 หลัก
                                                                       TextBoxVersionCode2.Text.Trim(), ' Version 2 หลัก
                                                                       TextBoxServiceCode3.Text.Trim(), ' Service 3 หลัก
                                                                       TextBoxPIDex.Text.Trim() ' 13 หลักที่ต้องการค้นหา
                                                                      )

                timeOut = 30
            Dim rec_pointer As IntPtr = Marshal.AllocHGlobal(reply_datamax)
            AMI.AMI_SESSION(SID, SID.Length)
            returnCode = AMI.AMI_REQUEST("", send5000, send5000.Length, rec_pointer, reply_datamax, reply_datasize, timeOut, status)
                Dim buf As String = TextBoxOutput.Text
                TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + Environment.NewLine

                Dim recive5000(reply_datasize) As Byte
                Marshal.Copy(rec_pointer, recive5000, 0, reply_datasize)
                Dim rc5000 = New AMI.Recive5000(recive5000)

                TextBoxOutput.Text = "Action: " + rc5000.ReActioCode4
                TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine
                TextBoxOutput.Text = TextBoxOutput.Text + "Office: " + TextBoxOfficeCode5.Text.Trim()
                TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine
                TextBoxOutput.Text = TextBoxOutput.Text + "Version: " + TextBoxVersionCode2.Text.Trim()
                TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine
                TextBoxOutput.Text = TextBoxOutput.Text + "Service: " + TextBoxServiceCode3.Text.Trim()
                TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine
                TextBoxOutput.Text = TextBoxOutput.Text + "PID: " + TextBoxPIDex.Text.Trim()
                TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine
                TextBoxOutput.Text = TextBoxOutput.Text + "ReturnCode: " + returnCode.ToString() + Environment.NewLine
                TextBoxOutput.Text = TextBoxOutput.Text + "Status: " + status.ToString() + " " + ami_stt.GetStatus3(status) + Environment.NewLine
                TextBoxOutput.Text = TextBoxOutput.Text + "ReturnStatus: " + rc5000.ReturnStatus5
                TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine
                TextBoxOutput.Text = TextBoxOutput.Text + "Data JSON: " + Environment.NewLine + rc5000.DataJSON + ">"
                TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + " " ' +buf

            'End If
        Catch ex As Exception
            Dim btn As Button = CType(sender, Button)
            MessageBox.Show(ex.Message,
            btn.Text,
            MessageBoxButtons.OK,
            MessageBoxIcon.Exclamation,
            MessageBoxDefaultButton.Button1)
        End Try

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If CheckBox1.Checked Then
            grantAccessLKbyPIN(sender, e)
        Else
            grantAccessLKbyCARD(sender, e)
        End If
    End Sub

    Private Sub grantAccessLKbyCARD(sender As Object, e As EventArgs)

        Try
            TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine
            returnCode = 99
            status = -99
            returnCode = SCAPI.OpenReader(ComboBox2.Text, status)

            If returnCode = 0 Then
                Dim atr As String = Space(100)
                Dim atr_len As Integer = 0
                Dim timeOut As Integer = 100
                Dim card_type As Integer = -999

                returnCode = 99
                status = -99
                returnCode = SCAPI.GetCardStatus(atr, atr_len, timeOut, card_type, status)

                'TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "5090 [" + returnCode.ToString() + "] " + status.ToString()
                'TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "5090 [ART] " + atr.Trim()

                Dim aid_bin(64) As Byte
                Dim aid_bin_len As Integer
                Dim util As New Utilities

                util.Str2Bin(SCAPI.MOI_AID, aid_bin, aid_bin_len)
                returnCode = 99
                status = -99
                returnCode = SCAPI.SelectApplet(aid_bin(0), aid_bin_len, status)

                If returnCode = 0 Then
                    'get PID
                    Dim block_id, offset As Integer
                    Dim Ppid As String
                    Dim data_size As Integer

                    block_id = 0
                    offset = 4
                    data_size = 13
                    Ppid = Space(13)

                    returnCode = 99
                    status = -99
                    returnCode = SCAPI.ReadData(block_id, offset, data_size, Ppid, status)

                    TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "pPID=" + Ppid
                    TextBoxPIDex.Text = Ppid

                    'อ่าน CID
                    Dim Pcid, pre_perso, perso, chip, os As String

                    Pcid = Space(16)
                    pre_perso = Space(20)
                    perso = Space(20)
                    chip = Space(20)
                    os = Space(20)

                    returnCode = 99
                    status = -99
                    returnCode = SCAPI.GetCardInfo(Pcid, chip, os, pre_perso, perso, status)

                    TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "pCID=" + Pcid

                    'Select ADM
                    util.Str2Bin(SCAPI.ADM_AID, aid_bin, aid_bin_len)
                    returnCode = 99
                    status = -99
                    returnCode = SCAPI.SelectApplet(aid_bin(0), aid_bin_len, status)

                    Dim adm_version, laser_number As String
                    Dim authorize, adm_state As Integer

                    adm_version = Space(5)
                    adm_state = 0
                    laser_number = Space(33)
                    authorize = 0

                    returnCode = 99
                    status = -99
                    returnCode = SCAPI.GetInfoADM(adm_version, adm_state, authorize, laser_number, status)

                    If authorize <> 0 Then
                        'GetRMAC
                        Dim xkey, rmacBk As String
                        Dim xkeyLen, rmacBkLen As Integer

                        'Dim xkey2 = BitConverter.ToString(Encoding.Default.GetBytes(rc9080.XKey32.Trim())).ToString().Replace("-", "")
                        xkey = BitConverter.ToString(Encoding.Default.GetBytes(rc9080.XKey32)).ToString().Replace("-", "")
                        xkeyLen = xkey.Length
                        rmacBk = Space(5000)
                        rmacBkLen = rmacBk.Length

                        returnCode = 99
                        status = -99
                        returnCode = SCAPI.GetRMAC(xkey, xkeyLen, rmacBk, rmacBkLen, status)
                        If returnCode <> 0 Then
                            TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "RMAC rc=" + returnCode.ToString() + " st=" + status.ToString() + "" '+ scapi_stt.GetStatus(1201L)
                        Else
                            TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "RMAC=" + rmacBkLen.ToString() + ":[" + rmacBk.Trim() + "]"

                            'EnvelopeGMSx
                            Dim envelop As String
                            Dim envelop_size As Integer

                            envelop_size = 255
                            envelop = Space(255)

                            returnCode = 99
                            status = -99
                            returnCode = EnvelopeGMSx(SCAPI.SAS_INT_AUTH_RMACKEY_ADMIN, rmacBk, rmacBkLen, envelop, envelop_size, status) 'Key SAS_INT_AUTH_RMACKEY_ADMIN

                            'AMI_REQUEST 5090
                            Dim reply_size As Integer
                            Dim reply_max As Integer

                            returnCode = 99
                            status = -99
                            timeOut = 1000

                            Dim send5090_2 As String = String.Format("{0}{1}{2}{3}{4}{5}:{6}",
                                                                   "5090",
                                                                   TextBoxTKey.Text.Trim(),
                                                                   Ppid.Trim(),
                                                                   Pcid.Trim(),
                                                                   "2",
                                                                   envelop_size,
                                                                   envelop.Substring(0, envelop_size)
                                                                  )

                            reply_size = 0
                            reply_max = 64 * 1024
                            Dim rec_pointer As IntPtr = Marshal.AllocHGlobal(reply_max)

                            returnCode = 99
                            status = -99
                            AMI.AMI_SESSION(SID, SID.Length)
                            returnCode = AMI.AMI_REQUEST("", send5090_2, send5090_2.Length, rec_pointer, reply_max, reply_size, timeOut, status)

                            If returnCode <> 0 Then
                                TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "5090 rc=" + returnCode.ToString() + " st=" + status.ToString() + " " + ami_stt.GetStatus3(status)
                            Else
                                Dim recive5090(reply_size) As Byte
                                Marshal.Copy(rec_pointer, recive5090, 0, reply_size)
                                Dim rc5090 = New AMI.Recive5090(recive5090)

                                If rc5090.ReturnStatus5 <> 0 Then
                                    TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "5090 " + rc5090.ReturnStatus5
                                Else
                                    TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "5090 Authorize Success By People CARD"
                                End If
                            End If
                        End If
                    End If
                Else

                End If
            Else
                TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "rc=" + returnCode.ToString() + " st=" + status.ToString()
            End If
        Catch ex As Exception
            Dim btn As Button = CType(sender, Button)
            MessageBox.Show(ex.Message,
            btn.Text,
            MessageBoxButtons.OK,
            MessageBoxIcon.Exclamation,
            MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Sub grantAccessLKbyPIN(sender As Object, e As EventArgs)
        returnCode = 0
        status = 0

        Try
            TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine
            returnCode = SCAPI.OpenReader(ComboBox2.Text, status)

            If returnCode = 0 Then
                returnCode = 0
                status = 0

                Dim atr As String = Space(100)
                Dim atr_len As Integer = 0
                Dim timeOut As Integer = 100
                Dim card_type As Integer = -999

                returnCode = SCAPI.GetCardStatus(atr, atr_len, timeOut, card_type, status)

                'TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "5090 [" + returnCode.ToString() + "] " + status.ToString()
                'TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "5090 [ART] " + atr.Trim()

                returnCode = 0
                status = 0

                Dim aid_bin(64) As Byte
                Dim aid_bin_len As Integer
                Dim util As New Utilities

                'select MOI
                util.Str2Bin(SCAPI.MOI_AID, aid_bin, aid_bin_len)
                returnCode = SCAPI.SelectApplet(aid_bin(0), aid_bin_len, status)

                If returnCode = 0 Then
                    returnCode = 0
                    status = 0
                    'get PID
                    Dim block_id, offset As Integer
                    Dim Ppid As String
                    Dim data_size As Integer

                    block_id = 0
                    offset = 4
                    data_size = 13
                    Ppid = Space(13)

                    returnCode = SCAPI.ReadData(block_id, offset, data_size, Ppid, status)
                    TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "pPID=" + Ppid
                    TextBoxPIDex.Text = Ppid

                    'อ่าน CID
                    Dim Pcid, pre_perso, perso, chip, os As String

                    Pcid = Space(16)
                    pre_perso = Space(20)
                    perso = Space(20)
                    chip = Space(20)
                    os = Space(20)
                    returnCode = 0
                    status = 0

                    returnCode = SCAPI.GetCardInfo(Pcid, chip, os, pre_perso, perso, status)
                    TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "pCID=" + Pcid

                    'Select ADM
                    util.Str2Bin(SCAPI.ADM_AID, aid_bin, aid_bin_len)
                    returnCode = SCAPI.SelectApplet(aid_bin(0), aid_bin_len, status)

                    'get Authorize
                    Dim adm_version, laser_number As String
                    Dim authorize, adm_state As Integer

                    adm_version = Space(5)
                    adm_state = 0
                    laser_number = Space(33)
                    authorize = 0

                    returnCode = SCAPI.GetInfoADM(adm_version, adm_state, authorize, laser_number, status)

                    If authorize = 0 Then
                        TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "5090 Cannot PIN Function"
                    Else
                        Dim try_remain As Integer
                        'PINCODE
                        returnCode = VerifyPIN(1, 0, try_remain, status)

                        If returnCode <> 0 Then
                            If status = 1001 Then
                                TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "5090 PIN Incorrect try = " + try_remain.ToString
                            End If
                        Else
                            'GetMatchStatus
                            Dim req_mode, req_type, match_status, random_size, cryto_size As Integer
                            Dim random, cryto As String

                            req_mode = 0
                            req_type = 1
                            random = util.Bin2Str(TextBoxKEY.Text, 32)
                            random_size = random.Length
                            cryto = Space(64)
                            cryto_size = 64
                            match_status = 0

                            returnCode = GetMatchStatus(req_type, req_mode, random, random_size, cryto, cryto_size, match_status, status)

                            If returnCode <> 0 Then
                                TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "5090 Get Match Status Error " + status.ToString
                            Else
                                If match_status <> 1 Then
                                    TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "5090 Match Status not match " + match_status.ToString
                                Else
                                    TextBox4.Text = cryto

                                    'EnvelopeGMSx
                                    Dim envelop As String
                                    Dim envelop_size As Integer

                                    envelop_size = 255
                                    envelop = Space(255)

                                    returnCode = EnvelopeGMSx(SCAPI.SAS_INT_AUTH_FPKEY_ADMIN, cryto, cryto_size, envelop, envelop_size, status) 'Key SAS_INT_AUTH_FPKEY_ADMIN

                                    If returnCode <> 0 Then
                                        TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "5090 Envelop not work"
                                    Else
                                        TextBox5.Text = envelop

                                        'AMI_REQUEST 5090
                                        Dim reply_size As Integer
                                        Dim reply_max As Integer

                                        returnCode = 0
                                        status = 0
                                        timeOut = 1000

                                        Dim send5090_1 As String = String.Format("{0}{1}{2}{3}{4}{5}:{6}",
                                                                   "5090",
                                                                   TextBoxTKey.Text.Trim(),
                                                                   Ppid.Trim(),
                                                                   Pcid.Trim(),
                                                                   "1",
                                                                   envelop_size,
                                                                   envelop.Substring(0, envelop_size)
                                                                  )

                                        reply_size = 0
                                        reply_max = 64 * 1024
                                        Dim rec_pointer As IntPtr = Marshal.AllocHGlobal(reply_max)

                                        returnCode = 99
                                        status = -99
                                        AMI.AMI_SESSION(SID, SID.Length)
                                        returnCode = AMI.AMI_REQUEST("", send5090_1, send5090_1.Length, rec_pointer, reply_max, reply_size, timeOut, status)

                                        If returnCode <> 0 Then
                                            TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "5090 [rc=" + returnCode.ToString() + "] [st=" + status.ToString() + "] " + ami_stt.GetStatus3(status)
                                        Else
                                            Dim recive5090(reply_size) As Byte
                                            Marshal.Copy(rec_pointer, recive5090, 0, reply_size)
                                            Dim rc5090 = New AMI.Recive5090(recive5090)
                                            If rc5090.ReturnStatus5 <> 0 Then
                                                TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "5090 [rs=" + rc5090.ReturnStatus5 + "] " + ami_stt.GetReturnStatus5(Long.Parse(rc5090.ReturnStatus5))
                                            Else
                                                TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "5090 Authorize Success By People PINCODE"
                                            End If
                                        End If

                                    End If
                                End If
                            End If
                        End If
                    End If
                Else

                End If
            Else
                TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + "rc=" + returnCode.ToString() + " st=" + status.ToString()
            End If
        Catch ex As Exception
            Dim btn As Button = CType(sender, Button)
            MessageBox.Show(ex.Message,
            btn.Text,
            MessageBoxButtons.OK,
            MessageBoxIcon.Exclamation,
            MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        'Dim DataRetun = requestCode5000()
        'Return

        Dim send_data As New AMI.Get5000
        Dim reply_data As New AMI.Reply5000


        Dim ReturnValue As Short
        Dim send_datasize As Integer
        Dim reply_datasize As Integer
        Dim reply_datamax As Integer
        Dim timeOut As Integer
        Dim status As Integer
        Try
            ReturnValue = 99
            status = 0

            reply_datasize = Marshal.SizeOf(reply_data)
            reply_datamax = Marshal.SizeOf(reply_data)

            send_data.Code = Encoding.Default.GetBytes("5000")
            send_data.XKey = Encoding.Default.GetBytes(TextBoxTKey.Text.Trim())

            send_data.OfficeCode = Encoding.Default.GetBytes("00023") ' สน.บท.
            send_data.VersionCode = Encoding.Default.GetBytes("01") ' Version 1
            send_data.ServiceCode = Encoding.Default.GetBytes("038") ' ภาพใบหน้า
            send_data.PID = Encoding.Default.GetBytes(TextBoxPIDex.Text.Trim())

            send_datasize = Marshal.SizeOf(send_data)
            timeOut = 30
            AMI.AMI_SESSION(SID, SID.Length)
            ReturnValue = AMI.AMI_REQUEST("", send_data, send_datasize, reply_data, reply_datamax, reply_datasize, timeOut, status)
            Dim buf As String = TextBoxOutput.Text
            TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + Environment.NewLine

            TextBoxOutput.Text = "Code " + Encoding.Default.GetString(reply_data.Reqno).Trim()
            'TextBoxOutput.Text = "Code " + DataRetun.Substring(0, 4)
            TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine
            TextBoxOutput.Text = TextBoxOutput.Text + "Office: 00023"
            TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine
            TextBoxOutput.Text = TextBoxOutput.Text + "Version: 01"
            TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine
            TextBoxOutput.Text = TextBoxOutput.Text + "Service: 038"
            TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine
            TextBoxOutput.Text = TextBoxOutput.Text + "PID: " + TextBoxPIDex.Text.Trim()
            TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine
            TextBoxOutput.Text = TextBoxOutput.Text + "ReturnValue: " + ReturnValue.ToString() + Environment.NewLine + "Status: "
            Dim stc As String = ""
            stc = scapi_stt.GetStatus(status)
            If stc.Substring(0, "Unknow Status".Length) = "Unknow Status" Then
                stc = ami_stt.GetStatus3(status)
            Else
                stc = scapi_stt.GetStatus(status)
            End If
            TextBoxOutput.Text = TextBoxOutput.Text + stc
            TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine
            TextBoxOutput.Text = TextBoxOutput.Text + "ReturnCode: " + Encoding.Default.GetString(reply_data.returnCode).Trim()
            'TextBoxOutput.Text = TextBoxOutput.Text + "ReturnCode " + DataRetun.Substring(4, 5)
            TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine
            TextBoxOutput.Text = TextBoxOutput.Text + "Data ----<" + Environment.NewLine + Encoding.UTF8.GetString(reply_data.Data).Trim() + ">"
            'TextBoxOutput.Text = TextBoxOutput.Text + "Data <" + DataRetun.Substring(9) + ">"
            TextBoxOutput.Text = TextBoxOutput.Text + Environment.NewLine + ">----" ' +buf

            showImage(Encoding.UTF8.GetString(reply_data.Data).Trim())

        Catch ex As Exception
            Dim btn As Button = CType(sender, Button)
            MessageBox.Show(ex.Message,
            btn.Text,
            MessageBoxButtons.OK,
            MessageBoxIcon.Exclamation,
            MessageBoxDefaultButton.Button1)
        End Try
    End Sub

    Private Sub showImage(jsonInput As String)
        'Throw New NotImplementedException()
        Dim frm As FormIMG = New FormIMG
        Dim img As LK_00023_01_038 = New LK_00023_01_038
        img = JsonConvert.DeserializeObject(Of LK_00023_01_038)(jsonInput)

        Dim imgg As Byte() = Convert.FromBase64String(img.image)
        Dim mm As New IO.MemoryStream(imgg)
        frm.PictureBox1.Image = Image.FromStream(mm)
        frm.Show()
        frm.Focus()
    End Sub
End Class
