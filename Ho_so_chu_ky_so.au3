#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <Word.au3>
#include <Excel.au3>
#include <MsgBoxConstants.au3>
#include <File.au3>
#include <Array.au3>

; --- CẤU HÌNH HỆ THỐNG ---
Global $oMyError = ObjEvent("AutoIt.Error", "_COMErrorFunc")
Global $sFilePathDoc = ""
Local $sFolderDanhSach = @ScriptDir & "\Danh_sach"
Local $sFolderLuuTru = @ScriptDir & "\Luu_tru"
Local $sExcelPath = $sFolderLuuTru & "\Danh_sach_ho_so_chu_ky_so.xlsx"

If Not FileExists($sFolderDanhSach) Then DirCreate($sFolderDanhSach)
If Not FileExists($sFolderLuuTru) Then DirCreate($sFolderLuuTru)

; --- GIAO DIỆN CHÍNH ---
Local $hGUI = GUICreate("HỆ THỐNG QUẢN LÝ HỒ SƠ CHỮ KÝ SỐ - PRO", 850, 750, -1, -1, $WS_OVERLAPPEDWINDOW)
GUISetFont(10, 400, 0, "Segoe UI")

; NHÓM CẤU HÌNH
GUICtrlCreateGroup(" CẤU HÌNH & TÌM KIẾM ", 10, 5, 830, 70)
Local $iBtnSelectDoc = GUICtrlCreateButton("Chọn file mẫu", 20, 30, 100, 30)
Local $lblDocPath = GUICtrlCreateLabel("Chưa chọn file mẫu...", 130, 37, 300, 20)
GUICtrlCreateLabel("MST cũ:", 440, 37, 50, 20)
Local $iSearchMST = GUICtrlCreateInput("", 500, 33, 160, 25)
Local $btnSearch = GUICtrlCreateButton("TÌM & ĐỔ DỮ LIỆU", 675, 30, 150, 30)
GUICtrlSetBkColor(-1, 0xFFF3CD)

; NHÓM I: DOANH NGHIỆP
GUICtrlCreateGroup(" I. THÔNG TIN DOANH NGHIỆP ", 10, 85, 830, 160)
GUICtrlCreateLabel("Tên công ty*:", 20, 110)
Local $iTenCty = GUICtrlCreateInput("", 120, 107, 430, 25)
GUICtrlCreateLabel("MST*:", 565, 110)
Local $iMST = GUICtrlCreateInput("", 620, 107, 210, 25)
GUICtrlCreateLabel("Ngày cấp phép:", 20, 142)
Local $iNgayCapP = GUICtrlCreateInput("", 120, 139, 150, 25)
GUICtrlCreateLabel("Nơi cấp phép:", 285, 142)
Local $iNoiCapP = GUICtrlCreateInput("", 375, 139, 455, 25)
GUICtrlCreateLabel("Địa chỉ ĐKKD*:", 20, 175)
Local $iDiaChiCty = GUICtrlCreateInput("", 120, 162, 710, 25)
GUICtrlCreateLabel("Email Cty:", 20, 207)
Local $iEmailCty = GUICtrlCreateInput("", 120, 204, 250, 25)
GUICtrlCreateLabel("SĐT Cty:", 385, 207)
Local $iSDTCty = GUICtrlCreateInput("", 455, 204, 150, 25)
GUICtrlCreateLabel("Tên viết tắt:", 615, 207)
Local $iTenVT = GUICtrlCreateInput("", 700, 204, 130, 25)

; NHÓM II: NGƯỜI ĐẠI DIỆN
GUICtrlCreateGroup(" II. NGƯỜI ĐẠI DIỆN PHÁP LUẬT ", 10, 255, 830, 175)
GUICtrlCreateLabel("Xưng hô:", 20, 280)
Local $iXungHo = GUICtrlCreateCombo("Ông", 90, 277, 70, 25)
GUICtrlSetData(-1, "Bà|Anh|Chị")
GUICtrlCreateLabel("Họ tên người ĐD*:", 170, 280)
Local $iNguoiDD = GUICtrlCreateInput("", 290, 277, 280, 25)
GUICtrlCreateLabel("Chức vụ*:", 580, 280)
Local $iChucVu = GUICtrlCreateInput("Giám đốc", 650, 277, 180, 25)
GUICtrlCreateLabel("CCCD/Hộ chiếu:", 20, 315)
Local $iCCCD = GUICtrlCreateInput("", 125, 312, 140, 25)
GUICtrlCreateLabel("Ngày cấp:", 275, 315)
Local $iNgayCapC = GUICtrlCreateInput("", 345, 312, 140, 25)
GUICtrlCreateLabel("Nơi cấp:", 495, 315)
Local $iNoiCapC = GUICtrlCreateInput("", 560, 312, 270, 25)
GUICtrlCreateLabel("Thường trú:", 20, 350)
Local $iHoKhau = GUICtrlCreateInput("", 100, 347, 730, 25)
GUICtrlCreateLabel("Email người ĐD:", 20, 385)
Local $iEmailNDD = GUICtrlCreateInput("", 120, 382, 290, 25)
GUICtrlCreateLabel("SĐT người ĐD:", 420, 385)
Local $iSDTNDD = GUICtrlCreateInput("", 540, 382, 290, 25)

; NHÓM III: DỊCH VỤ - ĐÃ CÂN CHỈNH TỌA ĐỘ Y
GUICtrlCreateGroup(" III. THÔNG TIN DỊCH VỤ & TÀI CHÍNH ", 10, 440, 830, 130)
GUICtrlCreateLabel("Loại hợp đồng:", 20, 465)
Local $iLoaiHD = GUICtrlCreateCombo("Cấp mới", 120, 462, 120, 25)
GUICtrlSetData(-1, "Gia hạn|Chuyển đổi")
GUICtrlCreateLabel("Số năm:", 250, 465)
Local $iSoNam = GUICtrlCreateInput("3", 310, 462, 50, 25)
GUICtrlCreateLabel("Số năm khác:", 370, 465)
Local $iSoNamKhac = GUICtrlCreateInput("", 465, 462, 80, 25)
GUICtrlCreateLabel("Serial/Mã BM:", 555, 465)
Local $iSerial = GUICtrlCreateInput("", 650, 462, 180, 25)
GUICtrlCreateLabel("Số TK ngân hàng:", 20, 500)
Local $iSTK = GUICtrlCreateInput("", 130, 497, 230, 25)
GUICtrlCreateLabel("Tại ngân hàng:", 370, 500)
Local $iNganHang = GUICtrlCreateInput("", 480, 497, 350, 25)
GUICtrlCreateLabel("Số tiền số:", 20, 535)
Local $iTienSo = GUICtrlCreateInput("", 100, 532, 200, 25)
GUICtrlCreateLabel("Số tiền chữ:", 310, 535)
Local $iTienChu = GUICtrlCreateInput("", 390, 532, 440, 25)

; NHÓM IV: ĐỊA DANH
GUICtrlCreateGroup(" IV. THỜI GIAN & ĐỊA DANH ", 10, 580, 830, 65)
GUICtrlCreateLabel("Địa danh:", 20, 607)
Local $iDiaDanh = GUICtrlCreateInput("Hà Nội", 90, 604, 150, 25)
GUICtrlCreateLabel("Ngày:", 260, 607)
Local $iNgay = GUICtrlCreateInput(@MDAY, 310, 604, 45, 25)
GUICtrlCreateLabel("Tháng:", 370, 607)
Local $iThang = GUICtrlCreateInput(@MON, 425, 604, 45, 25)
GUICtrlCreateLabel("Năm:", 485, 607)
Local $iNam = GUICtrlCreateInput(@YEAR, 530, 604, 60, 25)

Local $btnReset = GUICtrlCreateButton("NHẬP MỚI", 10, 670, 100, 45)
Local $btnPDF = GUICtrlCreateButton("XUẤT PDF & LƯU TẤT CẢ", 530, 670, 200, 45)
Local $btnPrint = GUICtrlCreateButton("IN HỒ SƠ", 740, 670, 100, 45)

; MẢNG INPUTS QUAN TRỌNG (B -> AD)
Global $aInputs = [$iTenCty, $iMST, $iNgayCapP, $iNoiCapP, $iDiaChiCty, $iEmailCty, $iSDTCty, $iTenVT, $iXungHo, $iNguoiDD, $iChucVu, $iCCCD, $iNgayCapC, $iNoiCapC, $iHoKhau, $iEmailNDD, $iSDTNDD, $iLoaiHD, $iSoNam, $iSoNamKhac, $iSerial, $iSTK, $iNganHang, $iTienSo, $iTienChu, $iDiaDanh, $iNgay, $iThang, $iNam]

GUISetState(@SW_SHOW)

; --- VÒNG LẶP SỰ KIỆN ---
While 1
    Local $nMsg = GUIGetMsg()
    Switch $nMsg
        Case $GUI_EVENT_CLOSE
            Exit
        Case $btnReset
            _ResetForm()
        Case $iBtnSelectDoc
            $sFilePathDoc = FileOpenDialog("Chọn file mẫu", @ScriptDir, "Word (*.docx)", 1)
            If Not @error Then GUICtrlSetData($lblDocPath, $sFilePathDoc)
        Case $btnSearch
            _SearchByExcelFind()
        Case $btnPDF, $btnPrint
            _HandleWorkflow($nMsg)
    EndSwitch
WEnd

; --- CÁC HÀM XỬ LÝ ---
Func _SearchByExcelFind()
    Local $sSearch = StringStripWS(GUICtrlRead($iSearchMST), 3)
    If $sSearch = "" Then Return MsgBox(48, "Thông báo", "Nhập MST cần tìm!")

    Local $oExcel = _Excel_Open(False)
    $oExcel.DisplayAlerts = False
    Local $oWB = _Excel_BookOpen($oExcel, $sExcelPath, True)
    If @error Then
        _Excel_Close($oExcel)
        Return MsgBox(16, "Lỗi", "Không tìm thấy file lưu trữ!")
    EndIf

    Local $oSheet = $oWB.ActiveSheet
    Local $oRangeFind = $oSheet.Columns("C").Find($sSearch)

    If IsObj($oRangeFind) Then
        Local $iRow = $oRangeFind.Row
        For $i = 0 To UBound($aInputs) - 1
            Local $sVal = $oSheet.Cells($iRow, $i + 2).Text ; Lấy Text chuẩn dd/mm/yyyy
            If $aInputs[$i] = $iXungHo Then
                GUICtrlSetData($iXungHo, "|Ông|Bà|Anh|Chị", $sVal) ; Thêm dấu | để reset list
            Else
                GUICtrlSetData($aInputs[$i], $sVal)
            EndIf
        Next
        _Excel_BookClose($oWB, False)
        _Excel_Close($oExcel, False)
        MsgBox(64, "Thành công", "Đã đổ dữ liệu cho MST: " & $sSearch)
    Else
        _Excel_BookClose($oWB, False)
        _Excel_Close($oExcel, False)
        MsgBox(48, "Thông báo", "Không tìm thấy MST: " & $sSearch)
    EndIf
EndFunc

Func _HandleWorkflow($iMsgID)
    If $sFilePathDoc == "" Then Return MsgBox(48, "Lỗi", "Hãy chọn file mẫu Word!")
    Local $sCurrMST = GUICtrlRead($iMST)
    If $sCurrMST == "" Then Return MsgBox(48, "Lỗi", "MST trống!")

    Local $aDataToSave[UBound($aInputs)]
    For $i = 0 To UBound($aInputs) - 1
        $aDataToSave[$i] = GUICtrlRead($aInputs[$i])
    Next
    _SaveToExcel($aDataToSave)

    Local $oWord = _Word_Create(), $oDoc = _Word_DocOpen($oWord, $sFilePathDoc, Default, Default, True)
    Local $aTags = ["ten_cong_ty", "ma_so_thue", "ngay_cap_phep", "noi_cap_phep", "dia_chi_cong_ty", "email_cong_ty", "dien_thoai_cong_ty", "ten_cong_ty_viet_tat", "xung_ho", "nguoi_dai_dien", "chuc_vu", "so_cmt", "ngay_cmt", "noi_cap_cmt", "ho_khau_thuong_tru", "email_nguoi_dai_dien", "dien_thoai_nguoi_dai_dien", "loai_hop_dong", "so_nam_cap_moi_gia_han", "so_nam_khac", "so_serial", "tai_khoan_so", "mo_tai_ngan_hang", "so_tien_bang_so", "so_tien_bang_chu", "dia_danh", "ng1", "thg1", "nam1"]

    For $i = 0 To UBound($aTags) - 1
        _Word_DocFindReplace($oDoc, "[[" & $aTags[$i] & "]]", GUICtrlRead($aInputs[$i]))
    Next

    Local $sNewDoc = $sFolderDanhSach & "\Ho_so_" & $sCurrMST & ".docx"
    Local $sNewPDF = $sFolderDanhSach & "\Ho_so_" & $sCurrMST & ".pdf"
    _Word_DocSaveAs($oDoc, $sNewDoc)

    If $iMsgID = $btnPrint Then
        $oWord.Visible = True
        $oDoc.PrintOut()
    Else
        _Word_DocExport($oDoc, $sNewPDF, 17)
        _Word_DocClose($oDoc)
        _Word_Quit($oWord)
        If FileExists($sNewPDF) Then ShellExecute($sNewPDF)
        ShellExecute($sFolderDanhSach)
    EndIf
EndFunc

Func _SaveToExcel($aDataRow)
    Local $oExcel = _Excel_Open(False)
    $oExcel.DisplayAlerts = False
    Local $oWB = FileExists($sExcelPath) ? _Excel_BookOpen($oExcel, $sExcelPath) : _Excel_BookNew($oExcel)
    Local $oSheet = $oWB.ActiveSheet
    Local $iLastRow = $oSheet.Range("C" & $oSheet.Rows.Count).End(-4162).Row
    Local $iNextRow = ($iLastRow < 3) ? 4 : $iLastRow + 1
    _Excel_RangeWrite($oWB, Default, $iNextRow - 3, "A" & $iNextRow)
    Local $aTable[1][UBound($aDataRow)]
    For $i = 0 To UBound($aDataRow) - 1
        $aTable[0][$i] = $aDataRow[$i]
    Next
    _Excel_RangeWrite($oWB, Default, $aTable, "B" & $iNextRow)
    _Excel_BookSave($oWB)
    _Excel_BookClose($oWB, False)
    _Excel_Close($oExcel, False)
EndFunc

Func _ResetForm()
    For $input In $aInputs
        GUICtrlSetData($input, "")
    Next
EndFunc

Func _COMErrorFunc()
EndFunc
