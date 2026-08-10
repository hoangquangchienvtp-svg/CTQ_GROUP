
function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Hàm giả lập tin nhắn Telegram gửi đến để kiểm tra lỗi
function testWebhookCall() {
    var props = PropertiesService.getScriptProperties();
    var chatId = props.getProperty("TELEGRAM_CHAT_ID");
    if (!chatId) {
        chatId = "d9189fc80494edcab485";
    } else {
        chatId = chatId.split(",")[0].trim();
    }

    var mockPayload = {
        "update_id": 123456789,
        "message": {
            "message_id": 123,
            "from": { "id": chatId, "is_bot": false, "first_name": "Test User" },
            "chat": { "id": chatId, "type": "private" },
            "date": Math.floor(Date.now() / 1000),
            "text": "1"
        }
    };

    var mockE = {
        "postData": {
            "contents": JSON.stringify(mockPayload)
        }
    };

    Logger.log("Bắt đầu chạy giả lập gửi lệnh '1'...");
    try {
        var result = doPost(mockE);
        var msg = "Chạy giả lập hoàn tất!\nKết quả trả về: " + (result ? result.getContent() : "null") + "\n\nBạn kiểm tra Telegram xem có nhận được báo cáo không nhé!";
        Logger.log(msg);
        SpreadsheetApp.getUi().alert(msg);
    } catch (e) {
        var errMsg = "Lỗi khi chạy giả lập: " + e.toString();
        Logger.log(errMsg);
        SpreadsheetApp.getUi().alert(errMsg);
    }
}

function incrementDataVersion() {
    var props = PropertiesService.getScriptProperties();
    var ver = parseInt(props.getProperty('GLOBAL_DATA_VERSION') || '0');
    props.setProperty('GLOBAL_DATA_VERSION', (ver + 1).toString());
}
function initSheets() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var userSheet = ss.getSheetByName("USERS");
    if (!userSheet) {
        userSheet = ss.insertSheet("USERS");
        userSheet.appendRow(["Username", "Password", "Role", "Permissions"]);
        userSheet.appendRow(["admin", "123", "admin", "dashboard,nhap-the,chinh-sua,in-the,nghi-viec,dulieu,diem-mu,cap-the,admin"]);
    }
    var adminDataSheet = ss.getSheetByName("ADMIN_DATA");
    if (!adminDataSheet) {
        adminDataSheet = ss.insertSheet("ADMIN_DATA");
        adminDataSheet.appendRow(["Tên Trạm", "Tên Line", "Tổ trưởng", "Bộ Phận", "Giao Viên", "Model", "Lý do"]);
    }
    var nhapTheSheet = ss.getSheetByName("NHẬP THẺ");
    if (!nhapTheSheet) {
        nhapTheSheet = ss.insertSheet("NHẬP THẺ");
        nhapTheSheet.appendRow(["Model", "Bộ Phận", "Mã Thẻ", "Họ Tên", "Giao Viên", "Ngày Vào", "Ngày Phát", "Loại Thẻ", "Đến Hạn", "CTQ 1", "CTQ 2", "CTQ 3"]);
    }
    var capTheSheet = ss.getSheetByName("CẤP THẺ");
    if (!capTheSheet) {
        capTheSheet = ss.insertSheet("CẤP THẺ");
        capTheSheet.appendRow(["BỘ PHẬN", "MODEL", "LINE", "MÃ THẺ", "HỌ TÊN", "LÝ DO CẤP LẠI", "TỔ TRƯỞNG YÊU CẦU", "CTQ 1"]);
    }
    var dataSheet = ss.getSheetByName("DATA");
    if (!dataSheet) {
        dataSheet = ss.insertSheet("DATA");
        dataSheet.appendRow(["TÊN TRẠM", "MÃ THẺ", "HỌ TÊN", "TIME", "LẦN TEST", "LỖI", "MÃ LỖI", "TÊN LỖI", "THỰC HIỆN", "LINE", "XÁC NHẬN", "NGÀY", "CA", "MODEL"]);
    }

    var lsCapTheSheet = ss.getSheetByName("LỊCH_SỬ_CẤP_THẺ");
    if (!lsCapTheSheet) {
        lsCapTheSheet = ss.insertSheet("LỊCH_SỬ_CẤP_THẺ");
        lsCapTheSheet.appendRow(["Thời Gian Yêu Cầu", "Bộ Phận", "Model", "Line", "Mã Thẻ", "Họ Tên", "Lý Do Xin Cấp", "Tổ Trưởng", "Trạm CTQ", "Thời Gian Xử Lý", "Người Xử Lý", "Trạng Thái", "Lý Do Từ Chối"]);
    }

    // KHỐI QUỸ LỢN ĐÃ ĐƯỢC ĐẶT CHUẨN VÀO BÊN TRONG HÀM
    var quyLonSheet = ss.getSheetByName("QUY_LON");
    if (!quyLonSheet) {
        quyLonSheet = ss.insertSheet("QUY_LON");
        quyLonSheet.appendRow(["NGÀY", "MÃ THẺ", "HỌ TÊN", "LỖI VI PHẠM", "ĐÃ NỘP (K)", "GHI NỢ (K)", "NGƯỜI THU"]);
    }

    // TỰ ĐỘNG TẠO SỔ QUỸ THU CHI NẾU CHƯA CÓ
    var quySheet = ss.getSheetByName("QUỸ");
    if (!quySheet) {
        quySheet = ss.insertSheet("QUỸ");
        quySheet.appendRow(["THỜI GIAN", "LOẠI", "SỐ TIỀN", "NGƯỜI NỘP/NHẬN", "LÝ DO", "NGƯỜI THU/CHI", "THÁNG GHI QUỸ"]);
    }


    var keHoachSheet = ss.getSheetByName("KE_HOACH");
    if (!keHoachSheet) {
        keHoachSheet = ss.insertSheet("KE_HOACH");
        keHoachSheet.appendRow(["CA", "MÃ THẺ", "HỌ TÊN", "GIÁO VIÊN", "T2", "T3", "T4", "T5", "T6", "T7", "CN"]);
    }

    // VS_NGAY sheet — Vệ Sinh Ca Ngày (TUẦN, STT, MÃ THẺ, HỌ TÊN, T2-T7)
    var vsNgaySheet = ss.getSheetByName("VS_NGAY");
    if (!vsNgaySheet) {
        vsNgaySheet = ss.insertSheet("VS_NGAY");
        vsNgaySheet.appendRow(["TUẦN", "STT", "MÃ THẺ", "HỌ TÊN", "T2", "T3", "T4", "T5", "T6", "T7"]);
    } else { _migrateAddTuanCol(vsNgaySheet); }

    var vsDemSheet = ss.getSheetByName("VS_DEM");
    if (!vsDemSheet) {
        vsDemSheet = ss.insertSheet("VS_DEM");
        vsDemSheet.appendRow(["TUẦN", "STT", "MÃ THẺ", "HỌ TÊN", "T2", "T3", "T4", "T5", "T6", "T7"]);
    } else { _migrateAddTuanCol(vsDemSheet); }

    // LT_NGAY sheet — Lịch Tuần Ca Ngày (TUẦN, STT, MÃ THẺ, HỌ TÊN, slot cols...)
    var ltNgaySheet = ss.getSheetByName("LT_NGAY");
    if (!ltNgaySheet) {
        ltNgaySheet = ss.insertSheet("LT_NGAY");
        ltNgaySheet.appendRow(["TUẦN", "STT", "MÃ THẺ", "HỌ TÊN", "T2_S1", "T2_S2", "T2_S3", "T3_S1", "T3_S2", "T3_S3", "T4_S1", "T4_S2", "T4_S3", "T5_S1", "T5_S2", "T5_S3", "T6_S1", "T6_S2", "T6_S3", "T7_S1", "T7_S2", "T7_S3"]);
    } else { _migrateAddTuanCol(ltNgaySheet); }

    var ltDemSheet = ss.getSheetByName("LT_DEM");
    if (!ltDemSheet) {
        ltDemSheet = ss.insertSheet("LT_DEM");
        ltDemSheet.appendRow(["TUẦN", "STT", "MÃ THẺ", "HỌ TÊN", "T2_S1", "T2_S2", "T2_S3", "T3_S1", "T3_S2", "T3_S3", "T4_S1", "T4_S2", "T4_S3", "T5_S1", "T5_S2", "T5_S3", "T6_S1", "T6_S2", "T6_S3", "T7_S1", "T7_S2", "T7_S3"]);
    } else { _migrateAddTuanCol(ltDemSheet); }


}

// ==========================================
// THÔNG TIN TRẠM — BACKEND FUNCTIONS
// ==========================================
function getThongTinTramData() {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("THONG_TIN_TRAM");
        if (!sheet) {
            sheet = ss.insertSheet("THONG_TIN_TRAM");
            sheet.appendRow(["Model", "Khu Vực", "Loại Trạm", "Tên Trạm", "Số Người Ngồi", "Ghi Chú", "Ngày Cập Nhật"]);
        }
        return sheet.getDataRange().getDisplayValues();
    } catch (e) {
        return [];
    }
}

function saveThongTinTram(rowData, currentUser, rowIndex) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(8000);
    } catch (e) {
        return { success: false, message: "Hệ thống đang bận. Vui lòng thử lại sau!" };
    }
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("THONG_TIN_TRAM");
        if (!sheet) {
            sheet = ss.insertSheet("THONG_TIN_TRAM");
            sheet.appendRow(["Model", "Khu Vực", "Loại Trạm", "Tên Trạm", "Số Người Ngồi", "Ghi Chú", "Ngày Cập Nhật"]);
            sheet.hideSheet();
        }
        var now = new Date().toLocaleString('vi-VN');
        var row = [
            rowData[0] || '',  // Model
            rowData[1] || '',  // Khu Vực
            rowData[2] || '',  // Loại Trạm
            rowData[3] || '',  // Tên Trạm
            rowData[4] || '',  // Số Người Ngồi
            rowData[5] || '',  // Ghi Chú
            now                // Ngày Cập Nhật
        ];

        if (rowIndex && rowIndex > 1) {
            sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
            SpreadsheetApp.flush();
            logAction("Đã Cập Nhật Trạm", currentUser, "Trạm: " + rowData[3] + " | Model: " + rowData[0] + " | Dòng: " + rowIndex);
        } else {
            sheet.appendRow(row);
            SpreadsheetApp.flush();
            logAction("Đã Thêm Trạm Mới", currentUser, "Trạm: " + rowData[3] + " | Model: " + rowData[0] + " | Khu Vực: " + rowData[1]);
        }
        return { success: true };
    } catch (e) {
        return { success: false, message: e.toString() };
    } finally {
        lock.releaseLock();
    }
}

function deleteThongTinTramRow(rowIndex, currentUser) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(8000);
    } catch (e) {
        return { success: false, message: "Hệ thống đang bận. Vui lòng thử lại sau!" };
    }
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("THONG_TIN_TRAM");
        if (!sheet) return { success: false, message: "Không tìm thấy sheet!" };
        var data = sheet.getDataRange().getDisplayValues();
        var detail = data[rowIndex - 1] ? (data[rowIndex - 1][3] + " (" + data[rowIndex - 1][0] + ")") : "Dòng " + rowIndex;
        sheet.deleteRow(rowIndex);
        SpreadsheetApp.flush();
        logAction("Đã Xóa Trạm", currentUser, "Trạm: " + detail);
        return { success: true };
    } catch (e) {
        return { success: false, message: e.toString() };
    } finally {
        lock.releaseLock();
    }
}

function doGet(e) {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        if (ss) {
            PropertiesService.getScriptProperties().setProperty("SPREADSHEET_ID", ss.getId());
        }
    } catch (err) { }
    initSheets();
    return HtmlService.createTemplateFromFile('CTQ2026_main').evaluate().setTitle('CTQ System Hub').addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

function moHeThongCTQ() {
    initSheets();
    var html = HtmlService.createTemplateFromFile('CTQ2026_main').evaluate().setTitle('Hệ Thống Quản Lý CTQ').setWidth(2600).setHeight(1850);
    SpreadsheetApp.getUi().showModalDialog(html, 'Hệ Thống Quản Lý CTQ');
}

function checkLogin(username, password) {
    // ✅ VALIDATION - NGĂN CHẶN BAD DATA
    if (!username || !password) return { success: false, message: "Vui lòng nhập tài khoản và mật khẩu" };
    if (String(username).trim().length < 3) return { success: false, message: "Tài khoản phải 3+ ký tự" };
    if (String(password).trim().length < 6) return { success: false, message: "Mật khẩu phải 6+ ký tự" };

    try {
        var props = PropertiesService.getScriptProperties();
        var user = username.toString().trim();
        var pass = password.toString().trim();

        // Kiểm tra tài khoản bị khóa
        var lockedStr = props.getProperty("LOCKED_USERS");
        if (lockedStr) {
            try {
                var locks = JSON.parse(lockedStr);
                if (locks[user]) {
                    if (locks[user].until > new Date().getTime()) {
                        var minsLeft = Math.ceil((locks[user].until - new Date().getTime()) / 60000);
                        var totalMins = locks[user].totalMinutes || minsLeft;
                        return { success: false, status: "locked", lockInfo: { reason: locks[user].reason, minsLeft: minsLeft, totalMinutes: totalMins, until: locks[user].until }, message: "Tài Khoản Đã Bị Khóa.\nLý do: " + locks[user].reason + ".\nMở khóa sau: " + minsLeft + " phút." };
                    } else {
                        delete locks[user];
                        props.setProperty("LOCKED_USERS", JSON.stringify(locks));
                    }
                }
            } catch (e) { }
        }

        // TỐI ƯU HÓA: Sử dụng CacheService để lưu thông tin USERS trong 5 phút
        var cache = CacheService.getScriptCache();
        var cachedUsersStr = cache.get("CACHED_USERS_DATA");
        var data;
        if (cachedUsersStr) {
            data = JSON.parse(cachedUsersStr);
        } else {
            var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USERS");
            if (sheet) {
                data = sheet.getDataRange().getDisplayValues();
                if (data.length > 1) {
                    try { cache.put("CACHED_USERS_DATA", JSON.stringify(data), 300); } catch (e) { }
                }
            } else {
                data = [];
            }
        }
        for (var i = 1; i < data.length; i++) {
            if (data[i][0] === user && data[i][1] === pass) {
                var sessionId = Math.random().toString(36).substring(2) + "_" + new Date().getTime();
                PropertiesService.getScriptProperties().setProperty("SESSION_ID_" + user, sessionId);
                // Đã chuyển recordUserLogin gọi ngầm từ client để login nhanh hơn
                return {
                    success: true,
                    role: data[i][2],
                    permissions: data[i][3] || "",
                    name: String(data[i][4] || user).substring(0, 50),
                    username: user,
                    sessionId: sessionId
                };
            }
        }
        return { success: false, message: 'Sai tài khoản hoặc mật khẩu!' };
    } catch (e) {
        return { success: false, message: 'Lỗi hệ thống: ' + e.toString() };
    }
}

function getInitialData() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var props = PropertiesService.getScriptProperties();
    var dataVersion = parseInt(props.getProperty('GLOBAL_DATA_VERSION') || '0');
    var fundHoldersStr = props.getProperty('FUND_HOLDERS');
    var fundHolders = fundHoldersStr ? JSON.parse(fundHoldersStr) : [];
    var dmtConfigStr = PropertiesService.getDocumentProperties().getProperty('DMT_CONFIG');
    var dmtConfig = dmtConfigStr ? JSON.parse(dmtConfigStr) : { "MAIN": [], "SUB": [] };

    var publicAnnSheet = ss.getSheetByName("PUBLIC_ANNOUNCEMENTS");
    var publicAnnData = '{}';
    if (publicAnnSheet) {
        var paValues = publicAnnSheet.getDataRange().getValues();
        var paJsonStr = "";
        for (var i = 0; i < paValues.length; i++) {
            if (paValues[i][0]) paJsonStr += paValues[i][0];
        }
        if (paJsonStr) publicAnnData = paJsonStr;
    } else {
        publicAnnData = props.getProperty('PUBLIC_ANNOUNCEMENTS') || '{}';
    }

    return {
        thongTinThe: (ss.getSheetByName("THÔNG TIN THẺ") ? ss.getSheetByName("THÔNG TIN THẺ").getDataRange().getDisplayValues() : []),
        adminData: ss.getSheetByName("ADMIN_DATA").getDataRange().getDisplayValues(),
        users: ss.getSheetByName("USERS").getDataRange().getDisplayValues(),
        nhapTheCount: (ss.getSheetByName("NHẬP THẺ") ? Math.max(0, ss.getSheetByName("NHẬP THẺ").getRange("A1:A").getValues().filter(String).length - 1) : 0),
        nhapTheData: (ss.getSheetByName("NHẬP THẺ") ? ss.getSheetByName("NHẬP THẺ").getDataRange().getDisplayValues() : []),
        capThe: (ss.getSheetByName("CẤP THẺ") ? ss.getSheetByName("CẤP THẺ").getDataRange().getDisplayValues() : []),
        hrData: (ss.getSheetByName("出勤数据") ? ss.getSheetByName("出勤数据").getDataRange().getDisplayValues() : []),
        sta5Data: (ss.getSheetByName("CHUYÊN CẦN STA5") ? ss.getSheetByName("CHUYÊN CẦN STA5").getDataRange().getDisplayValues() : []),
        quyLon: (ss.getSheetByName("QUY_LON") ? ss.getSheetByName("QUY_LON").getDataRange().getDisplayValues() : []),
        quy: (ss.getSheetByName("QUỸ") ? ss.getSheetByName("QUỸ").getDataRange().getDisplayValues() : []),
        fundHolders: fundHolders,
        dmtConfig: dmtConfig,
        bcTramConfig: (function () {
            var str = PropertiesService.getDocumentProperties().getProperty('BC_TRAM_CONFIG');
            return str ? JSON.parse(str) : [];
        })(),
        keHoachData: (ss.getSheetByName("KE_HOACH") ? ss.getSheetByName("KE_HOACH").getDataRange().getDisplayValues() : []),
        keHoachTitleDay: PropertiesService.getDocumentProperties().getProperty('KE_HOACH_TITLE_DAY') || 'Kế Hoạch Hỗ Trợ FATP (Ngày)',
        keHoachTitleNight: PropertiesService.getDocumentProperties().getProperty('KE_HOACH_TITLE_NIGHT') || 'Kế Hoạch Hỗ Trợ FATP (Đêm)',
        khStartDateDay: PropertiesService.getDocumentProperties().getProperty('KE_HOACH_DATE_DAY') || '',
        khStartDateNight: PropertiesService.getDocumentProperties().getProperty('KE_HOACH_DATE_NIGHT') || '',
        vsNgay: (ss.getSheetByName("VS_NGAY") ? ss.getSheetByName("VS_NGAY").getDataRange().getDisplayValues() : []),
        vsDem: (ss.getSheetByName("VS_DEM") ? ss.getSheetByName("VS_DEM").getDataRange().getDisplayValues() : []),
        ltNgay: (ss.getSheetByName("LT_NGAY") ? ss.getSheetByName("LT_NGAY").getDataRange().getDisplayValues() : []),
        ltDem: (ss.getSheetByName("LT_DEM") ? ss.getSheetByName("LT_DEM").getDataRange().getDisplayValues() : []),
        ltSlotConfig: _getLtSlotConfig(),
        marqueeText: PropertiesService.getDocumentProperties().getProperty('MARQUEE_TEXT') || '',
        marqueeExpiry: PropertiesService.getDocumentProperties().getProperty('MARQUEE_EXPIRY') || '',
        marqueeSpeed: PropertiesService.getDocumentProperties().getProperty('MARQUEE_SPEED') || '20',
        marqueeColor: PropertiesService.getDocumentProperties().getProperty('MARQUEE_COLOR') || '#ff0000',
        marqueeSize: PropertiesService.getDocumentProperties().getProperty('MARQUEE_SIZE') || '16',
        lockedUsers: props.getProperty('LOCKED_USERS') || '{}',
        publicAnnouncements: publicAnnData,
        cardApprovalConfig: getCardApprovalConfig(),
        dataVersion: dataVersion,
        loiThang7Data: (function () {
            try {
                var s = ss.getSheetByName("LỖI THÁNG 7");
                if (s) return s.getDataRange().getDisplayValues();
                var extSs = SpreadsheetApp.openById("1LjPv9qkzUAsfWvoV-3Vifr-xSltbT1ufJPc1iwD_c4k");
                var extS = extSs ? extSs.getSheetByName("LỖI THÁNG 7") : null;
                return extS ? extS.getDataRange().getDisplayValues() : [];
            } catch (e) {
                return [];
            }
        })(),
        chuyenCanStasData: (function () {
            try {
                var defaultLine = PropertiesService.getScriptProperties().getProperty('DEFAULT_STA_LINE') || 'STA5';
                var sheetName1 = 'CHUYÊN CẦN ' + defaultLine;
                var sheetName2 = 'CHUYÊN CẦN STAS';
                var s = ss.getSheetByName(sheetName1) || ss.getSheetByName(sheetName2);
                if (s) return s.getDataRange().getDisplayValues();
                var extSs = SpreadsheetApp.openById('1_T-OtIZF7_dVXZzTxtXzdLTr5KtWL642SD7kGk8fEH0');
                var extS = extSs ? (extSs.getSheetByName(sheetName1) || extSs.getSheetByName(sheetName2)) : null;
                return extS ? extS.getDataRange().getDisplayValues() : [];
            } catch (e) {
                return [];
            }
        })(),
        khongDungViTriData: (function () {
            try {
                var defaultLine = PropertiesService.getScriptProperties().getProperty('DEFAULT_STA_LINE') || 'STA5';
                var sheetName1 = 'KHÔNG ĐÚNG VỊ TRÍ ' + defaultLine;
                var sheetName2 = 'KHÔNG ĐÚNG VỊ TRÍ STAS';
                var sheetName3 = 'KHÔNG ĐÚNG VỊ TRÍ';
                var s = ss.getSheetByName(sheetName1) || ss.getSheetByName(sheetName2) || ss.getSheetByName(sheetName3);
                if (!s) {
                    var extSs = SpreadsheetApp.openById('1_T-OtIZF7_dVXZzTxtXzdLTr5KtWL642SD7kGk8fEH0');
                    s = extSs ? (extSs.getSheetByName(sheetName1) || extSs.getSheetByName(sheetName2) || extSs.getSheetByName(sheetName3)) : null;
                }
                if (s) {
                    var range = s.getDataRange();
                    var values = range.getDisplayValues();
                    var bgColors = range.getBackgrounds();
                    return { success: true, values: values, bgColors: bgColors };
                }
                return { success: false };
            } catch (e) {
                return { success: false, message: e.toString() };
            }
        })(),
        defaultStaLine: (PropertiesService.getScriptProperties().getProperty('DEFAULT_STA_LINE') || 'STA5')
    };
}

function getKhongDungViTriSta5Data() {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var s = ss.getSheetByName("KHÔNG ĐÚNG VỊ TRÍ STA5") || ss.getSheetByName("KHÔNG ĐÚNG VỊ TRÍ STAS") || ss.getSheetByName("KHÔNG ĐÚNG VỊ TRÍ");
        if (!s) {
            var extSs = SpreadsheetApp.openById("1_T-OtIZF7_dVXZzTxtXzdLTr5KtWL642SD7kGk8fEH0");
            if (extSs) {
                s = extSs.getSheetByName("KHÔNG ĐÚNG VỊ TRÍ STA5") || extSs.getSheetByName("KHÔNG ĐÚNG VỊ TRÍ STAS") || extSs.getSheetByName("KHÔNG ĐÚNG VỊ TRÍ");
                if (!s) {
                    var sheets = extSs.getSheets();
                    for (var i = 0; i < sheets.length; i++) {
                        var name = sheets[i].getName().toUpperCase();
                        if (name.indexOf("KHÔNG ĐÚNG VỊ TRÍ") !== -1 || name.indexOf("KHONG DUNG VI TRI") !== -1) {
                            s = sheets[i];
                            break;
                        }
                    }
                }
            }
        }
        if (s) {
            var range = s.getDataRange();
            var values = range.getDisplayValues();
            var bgColors = range.getBackgrounds();
            return { success: true, values: values, bgColors: bgColors };
        }
        return { success: false, message: "Không tìm thấy sheet KHÔNG ĐÚNG VỊ TRÍ STA5" };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}


// Hàm lấy danh sách tất cả sheet CHUYÊN CẦN và KHÔNG ĐÚNG VỊ TRÍ riêng (để điền 2 dropdown riêng)
function getAvailableStaLines() {
    try {
        var props = PropertiesService.getScriptProperties();
        var defCC = props.getProperty('DEFAULT_CC_SHEET') || 'CHUYÊN CẦN STA5';
        var defKDVT = props.getProperty('DEFAULT_KDVT_SHEET') || 'KHÔNG ĐÚNG VỊ TRÍ STA5';
        var prefix1 = 'CHUYÊN CẦN ';
        var prefix2 = 'KHÔNG ĐÚNG VỊ TRÍ ';

        var ccSheetsMap = {};
        var kdvtSheetsMap = {};

        var spreadsheets = [];
        try { spreadsheets.push(SpreadsheetApp.getActiveSpreadsheet()); } catch (e) { }
        try { spreadsheets.push(SpreadsheetApp.openById('1_T-OtIZF7_dVXZzTxtXzdLTr5KtWL642SD7kGk8fEH0')); } catch (e) { }

        spreadsheets.forEach(function (ssp) {
            if (!ssp) return;
            ssp.getSheets().forEach(function (sheet) {
                var name = sheet.getName().trim();
                var nameUp = name.toUpperCase();
                if (nameUp.indexOf(prefix1.toUpperCase()) === 0) {
                    ccSheetsMap[name] = true;
                } else if (nameUp.indexOf(prefix2.toUpperCase()) === 0) {
                    kdvtSheetsMap[name] = true;
                }
            });
        });

        var ccSheets = Object.keys(ccSheetsMap).map(function (n) { return { name: n }; });
        var kdvtSheets = Object.keys(kdvtSheetsMap).map(function (n) { return { name: n }; });

        // Đảm bảo luôn có STA5 fallback
        if (ccSheets.length === 0) ccSheets = [{ name: 'CHUYÊN CẦN STA5' }];
        if (kdvtSheets.length === 0) kdvtSheets = [{ name: 'KHÔNG ĐÚNG VỊ TRÍ STA5' }];

        return {
            success: true,
            ccSheets: ccSheets,
            kdvtSheets: kdvtSheets,
            defaultCCSheet: defCC,
            defaultKDVTSheet: defKDVT
        };
    } catch (e) {
        return {
            success: false, message: e.toString(),
            ccSheets: [{ name: 'CHUYÊN CẦN STA5' }],
            kdvtSheets: [{ name: 'KHÔNG ĐÚNG VỊ TRÍ STA5' }],
            defaultCCSheet: 'CHUYÊN CẦN STA5',
            defaultKDVTSheet: 'KHÔNG ĐÚNG VỊ TRÍ STA5'
        };
    }
}

// Hàm lưu 2 sheet mặc định riêng (CC + KDVT) vào Script Properties
function setDefaultStaSheets(ccSheetName, kdvtSheetName) {
    try {
        var props = PropertiesService.getScriptProperties();
        if (ccSheetName) props.setProperty('DEFAULT_CC_SHEET', ccSheetName.trim());
        if (kdvtSheetName) props.setProperty('DEFAULT_KDVT_SHEET', kdvtSheetName.trim());
        return { success: true };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

// Hàm load dữ liệu 1 sheet bất kỳ theo tên (dùng cho cả CC và KDVT)
function getSheetDataByName(sheetName, withBgColors) {
    try {
        sheetName = (sheetName || '').trim();
        var spreadsheets = [];
        try { spreadsheets.push(SpreadsheetApp.getActiveSpreadsheet()); } catch (e) { }
        try { spreadsheets.push(SpreadsheetApp.openById('1_T-OtIZF7_dVXZzTxtXzdLTr5KtWL642SD7kGk8fEH0')); } catch (e) { }

        var sheet = null;
        for (var i = 0; i < spreadsheets.length; i++) {
            if (!spreadsheets[i]) continue;
            sheet = spreadsheets[i].getSheetByName(sheetName);
            if (sheet) break;
        }

        if (!sheet) return { success: false, message: 'Không tìm thấy sheet: ' + sheetName };

        var range = sheet.getDataRange();
        var data = range.getDisplayValues();
        var bg = withBgColors ? range.getBackgrounds() : null;
        return { success: true, data: data, bgColors: bg };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

// Giữ nguyên hàm cũ để tương thích
function setDefaultStaLine(lineName) {
    return setDefaultStaSheets('CHUYÊN CẦN ' + (lineName || 'STA5').trim(), 'KHÔNG ĐÚNG VỊ TRÍ ' + (lineName || 'STA5').trim());
}
function getStaLineData(lineName) {
    lineName = (lineName || 'STA5').trim().toUpperCase();
    var cc = getSheetDataByName('CHUYÊN CẦN ' + lineName, false);
    var kdvt = getSheetDataByName('KHÔNG ĐÚNG VỊ TRÍ ' + lineName, true);
    return {
        success: true, line: lineName,
        chuyenCanData: cc.data || [],
        khongDungViTriData: kdvt.success ? { success: true, values: kdvt.data, bgColors: kdvt.bgColors } : { success: false }
    };
}
function getLoiThang7Data() {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("LỖI THÁNG 7");
        if (!sheet) {
            try {
                var extSs = SpreadsheetApp.openById("1LjPv9qkzUAsfWvoV-3Vifr-xSltbT1ufJPc1iwD_c4k");
                sheet = extSs ? extSs.getSheetByName("LỖI THÁNG 7") : null;
            } catch (e) { }
        }
        if (!sheet) return { success: false, data: [] };

        var range = sheet.getDataRange();
        var values = range.getDisplayValues();
        var richTexts = range.getRichTextValues();

        for (var i = 1; i < values.length; i++) {
            for (var j = 0; j < values[i].length; j++) {
                if (String(values[i][j]).trim().toLowerCase() === "link") {
                    var linkUrl = richTexts[i][j] ? richTexts[i][j].getLinkUrl() : null;
                    if (linkUrl) values[i][j] = linkUrl;
                }
            }
        }

        return { success: true, data: values };
    } catch (err) {
        return { success: false, message: err.toString(), data: [] };
    }
}

function getLoiWtData() {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("LỖI WT");
        if (!sheet) {
            try {
                // If it's stored in the external spreadsheet just like LỖI THÁNG 7
                var extSs = SpreadsheetApp.openById("1LjPv9qkzUAsfWvoV-3Vifr-xSltbT1ufJPc1iwD_c4k");
                sheet = extSs ? extSs.getSheetByName("LỖI WT") : null;
            } catch (e) { }
        }
        if (!sheet) return { success: false, data: [] };

        var range = sheet.getDataRange();
        var values = range.getDisplayValues();
        var richTexts = range.getRichTextValues();

        for (var i = 1; i < values.length; i++) {
            for (var j = 0; j < values[i].length; j++) {
                if (String(values[i][j]).trim().toLowerCase() === "link") {
                    var linkUrl = richTexts[i][j] ? richTexts[i][j].getLinkUrl() : null;
                    if (linkUrl) values[i][j] = linkUrl;
                }
            }
        }

        return { success: true, data: values };
    } catch (err) {
        return { success: false, message: err.toString(), data: [] };
    }
}
// XÓA NHẬP THẺ
function clearNhapTheData(type, rowNum) {

    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(6000);
    } catch (e) {
        return { success: false, message: "Hệ thống đang bận do có người khác đang thao tác. Vui lòng thử lại sau vài giây!" };
    }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NHẬP THẺ");
        if (!sheet) return { success: false, message: "Không tìm thấy sheet NHẬP THẺ" };
        if (type === 'ALL') { sheet.getRange("A2:L100").clearContent(); }
        else if (type === 'ROW' && rowNum >= 2) { sheet.deleteRow(rowNum); }
        var newCount = Math.max(0, sheet.getRange("A1:A").getValues().filter(String).length - 1);
        return { success: true, newCount: newCount };
    } catch (e) {
        return { success: false, message: e.toString() };
    } finally {
        lock.releaseLock();
    }
}

// LƯU ĐIỂM MÙ VÀO SHEET DATA
function saveDiemMuData(dataArr, currentUser) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
    } catch (e) {
        return { success: false, message: "Hệ thống đang bận do có người khác đang thao tác. Vui lòng thử lại sau vài giây!" };
    }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DATA');
        if (!sheet) return { success: false, message: "Không tìm thấy sheet DATA" };

        var colBValues = sheet.getRange('B:B').getDisplayValues().flat();
        var lastRow = 0;
        for (var r = colBValues.length - 1; r >= 0; r--) {
            if (colBValues[r] !== "") { lastRow = r + 1; break; }
        }
        if (lastRow < 1) lastRow = 1;

        sheet.getRange(lastRow + 1, 1, dataArr.length, dataArr[0].length).setValues(dataArr);
        var line = dataArr[0] && dataArr[0][9] ? dataArr[0][9] : "N/A";
        var toTruong = dataArr[0] && dataArr[0][10] ? dataArr[0][10] : "N/A";
        var ca = dataArr[0] && dataArr[0][12] ? dataArr[0][12] : "N/A";

        var dsLoi = [];
        for (var i = 0; i < dataArr.length; i++) {
            var soLoi = parseInt(dataArr[i][5]) || 0;
            if (soLoi > 0) {
                var maLoi = dataArr[i][6] || "";
                var tenLoi = dataArr[i][7] || "";
                var nguoiThucHien = dataArr[i][8] || "";
                var id = dataArr[i][1] || "";
                var hoTen = dataArr[i][2] || "";
                var thoiGian = dataArr[i][3] || "";
                var lanTest = dataArr[i][4] || "";
                var tenTram = dataArr[i][0] || "";
                var ngay = dataArr[i][11] || "";

                dsLoi.push("- " + maLoi + " (" + tenLoi + "): " + soLoi + " lỗi, do " + nguoiThucHien + " nhập.");

                // Gửi thông báo Telegram
                var telegramMsg = "[CTQ] 🚨 PHÁT SINH LỖI ĐIỂM MF\n";
                telegramMsg += "Người thao tác: " + (currentUser || "Khách") + "\n";
                telegramMsg += "Chi tiết:\n";
                telegramMsg += "Line: " + line + " | Ca: " + ca + "\n";
                telegramMsg += "Tổ trưởng: " + toTruong + "\n";
                telegramMsg += "ID: " + id + " / Họ Tên: " + hoTen + "\n";
                telegramMsg += "Chi tiết lỗi:\n";
                telegramMsg += "- Tên Trạm: " + tenTram + "\n";
                telegramMsg += "- Mã Lỗi: " + maLoi + "\n";
                telegramMsg += "- Tên Lỗi: " + tenLoi + "\n";
                telegramMsg += "- Lần Test: " + lanTest + "\n";
                telegramMsg += "- Số Lỗi: " + soLoi + " lỗi\n";
                telegramMsg += "- Người Nhập: " + nguoiThucHien + "\n";
                telegramMsg += "Thời gian: " + thoiGian + " " + ngay;

                sendTelegramNotify(telegramMsg);
            }
        }

        if (dsLoi.length > 0) {
            logAction("🚨 PHÁT SINH LỖI MỚI", currentUser, "Line: " + line + " | Ca: " + ca + "\nTổ trưởng: " + toTruong + "\n\nChi tiết lỗi:\n" + dsLoi.join("\n"));
        } else {
            logAction("Thêm Điểm Mù (Không lỗi)", currentUser, "Line: " + line + " | Ca: " + ca + "\nTổ trưởng: " + toTruong + "\nSố lượng nhập: " + dataArr.length + " dòng (Tất cả đều đạt).");
        }
        return { success: true };
    } catch (e) {
        return { success: false, message: "Lỗi hệ thống: " + e.toString() };
    } finally {
        lock.releaseLock();
    }
}

function getPhanTichData(fromDate, toDate) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dataSheet = ss.getSheetByName("DATA");
    if (!dataSheet) return [];
    var lastRow = dataSheet.getLastRow();
    if (lastRow < 2) return [];
    var data = dataSheet.getRange(2, 1, lastRow - 1, 13).getDisplayValues();

    var fDate, tDate;
    if (fromDate && toDate) {
        var dp1 = fromDate.split('-'); fDate = new Date(dp1[0], dp1[1] - 1, dp1[2]); fDate.setHours(0, 0, 0, 0);
        var dp2 = toDate.split('-'); tDate = new Date(dp2[0], dp2[1] - 1, dp2[2]); tDate.setHours(23, 59, 59, 999);
    }

    var agg = {};
    for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var rowDateStr = row[11];
        if (!rowDateStr) continue;

        var rDate;
        var parts = String(rowDateStr).split('/');
        if (parts.length === 3) {
            rDate = new Date(parts[2], parts[1] - 1, parts[0]);
        } else {
            rDate = new Date(rowDateStr);
        }

        if (fDate && tDate) {
            if (isNaN(rDate.getTime()) || rDate < fDate || rDate > tDate) continue;
        }

        var dStr = ('0' + rDate.getDate()).slice(-2) + '/' + ('0' + (rDate.getMonth() + 1)).slice(-2);
        var kiemTra = parseFloat(row[4]) || 0;
        var loi = parseFloat(row[5]) || 0;

        if (!agg[dStr]) agg[dStr] = { kiemTra: 0, loi: 0, sortVal: rDate.getTime() };
        agg[dStr].kiemTra += kiemTra;
        agg[dStr].loi += loi;
    }

    var resultArr = [];
    for (var key in agg) {
        var rate = agg[key].kiemTra > 0 ? (agg[key].loi / agg[key].kiemTra) * 100 : 0;
        resultArr.push({
            date: key,
            kiemTra: agg[key].kiemTra,
            loi: agg[key].loi,
            rate: rate,
            sortVal: agg[key].sortVal
        });
    }
    resultArr.sort(function (a, b) { return a.sortVal - b.sortVal; });
    return resultArr;
}

function saveDMTConfig(configObj) {
    PropertiesService.getDocumentProperties().setProperty('DMT_CONFIG', JSON.stringify(configObj));
    return { success: true };
}

function saveBcTramConfig(configObj) {
    PropertiesService.getDocumentProperties().setProperty('BC_TRAM_CONFIG', JSON.stringify(configObj));
    return { success: true };
}

function addMaLoi(code, name, currentUser) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
    } catch (e) {
        throw new Error("Hệ thống đang bận do có người khác đang thao tác. Vui lòng thử lại sau vài giây!");
    }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ADMIN_DATA');
        var data = sheet.getDataRange().getValues();
        var rowToInsert = 2;
        while (rowToInsert <= data.length && data[rowToInsert - 1][8] !== "") { rowToInsert++; }
        sheet.getRange(rowToInsert, 9).setValue(code);
        sheet.getRange(rowToInsert, 10).setValue(name);
        logAction("Thêm Mã Lỗi", currentUser, "Mã: " + code + " - Tên: " + name);
    } finally {
        lock.releaseLock();
    }
}

function xoaNhanVienNghiViec(rowsToDelete, currentUser) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
    } catch (e) {
        throw new Error("Hệ thống đang bận do có người khác đang thao tác. Vui lòng thử lại sau vài giây!");
    }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('THÔNG TIN THẺ');
        var data = sheet.getDataRange().getDisplayValues();
        var deletedDetails = [];
        for (var i = 0; i < rowsToDelete.length; i++) {
            var r = rowsToDelete[i] - 1; // 0-indexed data array
            if (data[r]) {
                deletedDetails.push("Mã Thẻ: " + data[r][3] + " - Tên: " + data[r][4]);
            }
        }
        rowsToDelete.sort(function (a, b) { return b - a });
        for (var i = 0; i < rowsToDelete.length; i++) { sheet.deleteRow(rowsToDelete[i]); }
        SpreadsheetApp.flush();
        logAction("Xoá Nhân Viên Nghỉ Việc", currentUser, "Đã xoá " + rowsToDelete.length + " nhân viên:\n" + deletedDetails.join("\n"));
        return getInitialData().thongTinThe;
    } finally {
        lock.releaseLock();
    }
}

function saveOrUpdateUser(username, password, role, permissions, currentUser) {
    incrementDataVersion();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USERS");
    var data = sheet.getDataRange().getDisplayValues();
    for (var i = 1; i < data.length; i++) {
        if (data[i][0] === username) {
            if (password) sheet.getRange(i + 1, 2).setValue(password);
            sheet.getRange(i + 1, 3).setValue(role);
            sheet.getRange(i + 1, 4).setValue(permissions);
            logAction("Cập Nhật User", currentUser, "User: " + username + " đã được cập nhật.");
            return true;
        }
    }
    if (!password) password = "123"; // Default password if new user and not provided
    sheet.appendRow([username, password, role, permissions, "online", ""]);
    logAction("Thêm User Mới", currentUser, "User: " + username);
    return true;
}

function saveMultipleUsers(usernames, password, role, permissions, currentUser) {
    incrementDataVersion();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USERS");
    var data = sheet.getDataRange().getDisplayValues();
    var props = PropertiesService.getScriptProperties();

    for (var u = 0; u < usernames.length; u++) {
        var username = usernames[u];
        var found = false;
        for (var i = 1; i < data.length; i++) {
            if (data[i][0] === username) {
                var oldPerms = data[i][3] || "";
                if (password) sheet.getRange(i + 1, 2).setValue(password);
                sheet.getRange(i + 1, 3).setValue(role);
                sheet.getRange(i + 1, 4).setValue(permissions);

                var oldList = oldPerms.split(",").map(function (s) { return s.trim(); }).filter(function (s) { return s; });
                var newList = permissions.split(",").map(function (s) { return s.trim(); }).filter(function (s) { return s; });

                var added = [];
                for (var j = 0; j < newList.length; j++) {
                    if (oldList.indexOf(newList[j]) === -1) added.push(newList[j]);
                }

                if (added.length > 0) {
                    var existingKey = "NEW_PERMS_" + username.toUpperCase();
                    var existingJson = props.getProperty(existingKey);
                    var existingArr = [];
                    try { if (existingJson) existingArr = JSON.parse(existingJson); } catch (e) { }

                    for (var a = 0; a < added.length; a++) {
                        if (existingArr.indexOf(added[a]) === -1) {
                            existingArr.push(added[a]);
                        }
                    }
                    props.setProperty(existingKey, JSON.stringify(existingArr));

                    var actionMsg = "Đã thêm " + added.length + " quyền mới cho User " + username + ".\n" +
                        "Tổng cộng User " + username + " đang có " + newList.length + " quyền như sau:\n" + newList.join(", ");
                    logAction("Cập Nhật Quyền User", currentUser, actionMsg);
                } else if (oldList.length !== newList.length || oldList.join(",") !== newList.join(",")) {
                    logAction("Cập Nhật Quyền User", currentUser, "Cập nhật quyền cho User " + username + ": " + newList.length + " quyền.");
                } else {
                    logAction("Cập Nhật User", currentUser, "Cập nhật thông tin User: " + username);
                }
                found = true;
                break;
            }
        }
        if (!found) {
            var pwd = password ? password : "123";
            sheet.appendRow([username, pwd, role, permissions, "online", ""]);
            var addedList = permissions.split(",").map(function (s) { return s.trim(); }).filter(function (s) { return s; });
            if (addedList.length > 0) {
                props.setProperty("NEW_PERMS_" + username.toUpperCase(), JSON.stringify(addedList));
            }
            logAction("Thêm User Mới", currentUser, "Đã thêm User " + username + ".\nTổng cộng có " + addedList.length + " quyền như sau:\n" + permissions);
        }
    }
    SpreadsheetApp.flush();
    return { success: true, message: "Đã xử lý " + usernames.length + " users." };
}

function checkAndGetNewPerms(username) {
    try {
        if (!username) return { success: true, newPerms: [] };
        var props = PropertiesService.getScriptProperties();
        var key = "NEW_PERMS_" + String(username).trim().toUpperCase();
        var val = props.getProperty(key);
        if (!val) return { success: true, newPerms: [] };

        var arr = JSON.parse(val);
        return { success: true, newPerms: arr };
    } catch (e) {
        return { success: false, newPerms: [] };
    }
}

function clearUserNewPerms(username) {
    try {
        if (!username) return { success: true };
        var props = PropertiesService.getScriptProperties();
        var key = "NEW_PERMS_" + String(username).trim().toUpperCase();
        props.deleteProperty(key);
        return { success: true };
    } catch (e) {
        return { success: false };
    }
}

function deleteUser(username, currentUser) {
    incrementDataVersion();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USERS");
    var data = sheet.getDataRange().getDisplayValues();
    for (var i = 1; i < data.length; i++) {
        if (data[i][0] === username) {
            sheet.deleteRow(i + 1);
            logAction("Xoá User", currentUser, "User: " + username + " đã bị xoá khỏi hệ thống.");
            return true;
        }
    } return false;
}

function saveMarqueeSettings(text, expiry, speed, color, size) {
    try {
        var props = PropertiesService.getScriptProperties();
        props.setProperty("MARQUEE_TEXT", text || "");
        props.setProperty("MARQUEE_EXPIRY", expiry || "");
        props.setProperty("MARQUEE_SPEED", speed || "20");
        props.setProperty("MARQUEE_COLOR", color || "#ff0000");
        props.setProperty("MARQUEE_SIZE", size || "16");
        return { success: true };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function savePublicAnnouncements(data) {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("PUBLIC_ANNOUNCEMENTS");
        if (!sheet) {
            sheet = ss.insertSheet("PUBLIC_ANNOUNCEMENTS");
            sheet.hideSheet();
        }

        // Tự động dịch sang tiếng Trung và tiếng Anh (nếu chưa nhập tay hoặc số dòng không khớp)
        for (var i = 1; i <= 4; i++) {
            if (data['ann' + i]) {
                if (!data['ann' + i + '_cn']) { try { data['ann' + i + '_cn'] = LanguageApp.translate(data['ann' + i], '', 'zh-CN'); } catch (e) { } }
                if (!data['ann' + i + '_en']) { try { data['ann' + i + '_en'] = LanguageApp.translate(data['ann' + i], '', 'en'); } catch (e) { } }
            }
            if (data['customList' + i]) {
                var vnLines = data['customList' + i].split(/\r?\n/).filter(function (l) { return l.trim().length > 0; }).length;
                var cnLines = (data['customList' + i + '_cn'] || '').split(/\r?\n/).filter(function (l) { return l.trim().length > 0; }).length;
                var enLines = (data['customList' + i + '_en'] || '').split(/\r?\n/).filter(function (l) { return l.trim().length > 0; }).length;

                if (!data['customList' + i + '_cn'] || vnLines !== cnLines) {
                    try {
                        data['customList' + i + '_cn'] = data['customList' + i].split(/\r?\n/).map(function (line) {
                            return line.trim() ? LanguageApp.translate(line, '', 'zh-CN') : '';
                        }).join('\n');
                    } catch (e) { }
                }
                if (!data['customList' + i + '_en'] || vnLines !== enLines) {
                    try {
                        data['customList' + i + '_en'] = data['customList' + i].split(/\r?\n/).map(function (line) {
                            return line.trim() ? LanguageApp.translate(line, '', 'en') : '';
                        }).join('\n');
                    } catch (e) { }
                }
            }
        }

        var jsonStr = JSON.stringify(data);
        var chunks = [];
        for (var i = 0; i < jsonStr.length; i += 45000) {
            chunks.push([jsonStr.substring(i, i + 45000)]);
        }
        sheet.clear();
        if (chunks.length > 0) {
            sheet.getRange(1, 1, chunks.length, 1).setValues(chunks);
        }
        return { success: true };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function getPublicAnnouncements() {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("PUBLIC_ANNOUNCEMENTS");
        if (!sheet) {
            return { success: true, data: PropertiesService.getScriptProperties().getProperty("PUBLIC_ANNOUNCEMENTS") || "{}" };
        }
        var values = sheet.getDataRange().getValues();
        var jsonStr = "";
        for (var i = 0; i < values.length; i++) {
            if (values[i][0]) jsonStr += values[i][0];
        }
        return { success: true, data: jsonStr || "{}" };
    } catch (e) {
        return { success: false };
    }
}

// === QUẢN LÝ BẢNG TIN ===
function saveBangTinConfig(configData, currentUser) {
    try {
        var props = PropertiesService.getScriptProperties();
        props.setProperty("BANG_TIN_CONFIG", JSON.stringify(configData));
        logAction("Lưu Bảng Tin", currentUser, "Đã cập nhật nội dung Bảng Tin");
        return { success: true };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function getBangTinConfig() {
    try {
        var props = PropertiesService.getScriptProperties();
        var data = props.getProperty("BANG_TIN_CONFIG") || "{}";
        return { success: true, data: data };
    } catch (e) {
        return { success: false };
    }
}

function lockUserAccount(username, reason, minutes, currentUser) {
    try {
        var props = PropertiesService.getScriptProperties();
        var lockedStr = props.getProperty("LOCKED_USERS") || "{}";
        var locks = JSON.parse(lockedStr);
        locks[username] = {
            reason: reason,
            until: new Date().getTime() + (minutes * 60000),
            totalMinutes: minutes
        };
        props.setProperty("LOCKED_USERS", JSON.stringify(locks));
        logAction("Khóa User", currentUser, "Khóa user: " + username + " trong " + minutes + " phút. Lý do: " + reason);
        return { success: true };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function unlockUserAccount(username, currentUser) {
    try {
        var props = PropertiesService.getScriptProperties();
        var lockedStr = props.getProperty("LOCKED_USERS") || "{}";
        var locks = JSON.parse(lockedStr);
        if (locks[username]) {
            delete locks[username];
            props.setProperty("LOCKED_USERS", JSON.stringify(locks));
            logAction("Mở Khóa User", currentUser, "Mở khóa user: " + username);
        }
        return { success: true };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

// === BATCH OPERATIONS CHO ADMIN ===
function deleteMultipleUsers(usernames, currentUser) {
    var lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (e) {
        return { success: false, message: "Hệ thống đang bận. Vui lòng thử lại!" };
    }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USERS");
        var data = sheet.getDataRange().getDisplayValues();
        var rowsToDelete = [];
        for (var i = 1; i < data.length; i++) {
            if (usernames.indexOf(data[i][0]) !== -1) {
                rowsToDelete.push(i + 1);
            }
        }
        rowsToDelete.sort(function (a, b) { return b - a; });
        for (var j = 0; j < rowsToDelete.length; j++) {
            sheet.deleteRow(rowsToDelete[j]);
        }
        SpreadsheetApp.flush();
        logAction("Xoá Nhiều User", currentUser, "Đã xoá " + usernames.length + " user: " + usernames.join(", "));
        return { success: true, message: "Đã xoá " + rowsToDelete.length + " tài khoản!" };
    } catch (e) {
        return { success: false, message: e.toString() };
    } finally {
        lock.releaseLock();
    }
}

function lockMultipleUsers(usernames, reason, minutes, currentUser) {
    try {
        var props = PropertiesService.getScriptProperties();
        var lockedStr = props.getProperty("LOCKED_USERS") || "{}";
        var locks = JSON.parse(lockedStr);
        var until = new Date().getTime() + (minutes * 60000);
        for (var i = 0; i < usernames.length; i++) {
            locks[usernames[i]] = { reason: reason, until: until, totalMinutes: minutes };
        }
        props.setProperty("LOCKED_USERS", JSON.stringify(locks));
        logAction("Khóa Nhiều User", currentUser, "Khóa " + usernames.length + " user trong " + minutes + " phút. Lý do: " + reason + ". Users: " + usernames.join(", "));
        return { success: true, message: "Đã khóa " + usernames.length + " tài khoản!" };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function unlockMultipleUsers(usernames, currentUser) {
    try {
        var props = PropertiesService.getScriptProperties();
        var lockedStr = props.getProperty("LOCKED_USERS") || "{}";
        var locks = JSON.parse(lockedStr);
        var count = 0;
        for (var i = 0; i < usernames.length; i++) {
            if (locks[usernames[i]]) { delete locks[usernames[i]]; count++; }
        }
        props.setProperty("LOCKED_USERS", JSON.stringify(locks));
        logAction("Mở Khóa Nhiều User", currentUser, "Mở khóa " + count + " user: " + usernames.join(", "));
        return { success: true, message: "Đã mở khóa " + count + " tài khoản!" };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function updatePermissionsMultiple(usernames, permissions, role, currentUser) {
    var lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (e) {
        return { success: false, message: "Hệ thống đang bận. Vui lòng thử lại!" };
    }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USERS");
        var data = sheet.getDataRange().getDisplayValues();
        var count = 0;
        for (var i = 1; i < data.length; i++) {
            if (usernames.indexOf(data[i][0]) !== -1) {
                sheet.getRange(i + 1, 3).setValue(role);
                sheet.getRange(i + 1, 4).setValue(permissions);
                count++;
            }
        }
        SpreadsheetApp.flush();
        logAction("Cập Nhật Quyền Nhiều User", currentUser, "Cập nhật quyền cho " + count + " user: " + usernames.join(", ") + "\nQuyền mới: " + permissions);
        return { success: true, message: "Đã cập nhật quyền cho " + count + " tài khoản!" };
    } catch (e) {
        return { success: false, message: e.toString() };
    } finally {
        lock.releaseLock();
    }
}

function getAdminColName(colIndex) {
    var names = {
        0: "Tên Trạm", 1: "Tên Line", 2: "Tổ trưởng", 3: "Bộ Phận",
        4: "Giáo Viên", 5: "Model", 6: "Lý do", 7: "Lý do vi phạm",
        16: "IPQC", 17: "ME", 18: "CTQ GIÁO VIÊN"
    };
    return names[colIndex] || "Cột " + (colIndex + 1);
}

function addAdminData(colIndex, value, currentUser) {
    incrementDataVersion();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ADMIN_DATA");
    var data = sheet.getDataRange().getValues();
    var rowToInsert = 2;
    while (rowToInsert <= data.length && data[rowToInsert - 1][colIndex] !== "") { rowToInsert++; }
    sheet.getRange(rowToInsert, colIndex + 1).setValue(value);
    logAction("Thêm Dữ Liệu ADMIN_DATA", currentUser, "Cột: " + getAdminColName(colIndex) + " - Giá trị: " + value);
    return sheet.getDataRange().getDisplayValues();
}

function deleteAdminData(colIndex, value, currentUser) {
    incrementDataVersion();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ADMIN_DATA");
    var data = sheet.getDataRange().getValues();
    var newCol = [];
    for (var i = 1; i < data.length; i++) { if (data[i][colIndex].toString().trim() !== value.toString().trim() && data[i][colIndex] !== "") newCol.push([data[i][colIndex]]); }
    newCol.push([""]);
    sheet.getRange(2, colIndex + 1, data.length, 1).clearContent();
    if (newCol.length > 0 && newCol[0][0] !== undefined) sheet.getRange(2, colIndex + 1, newCol.length, 1).setValues(newCol);
    logAction("Xoá Dữ Liệu ADMIN_DATA", currentUser, "Cột: " + getAdminColName(colIndex) + " - Giá trị: " + value);
    return sheet.getDataRange().getDisplayValues();
}

function getEmployeeData(empId) {
    try {
        if (!empId) return { success: false, name: "", date: "" };
        var searchCode = String(empId).trim().toUpperCase();

        var spreadSheets = [];
        try {
            var activeSs = SpreadsheetApp.getActiveSpreadsheet();
            if (activeSs) spreadSheets.push(activeSs);
        } catch (eAct) { }

        var extraIds = [
            "1_T-OtIZF7_dVXZzTxtXzdLTr5KtWL642SD7kGk8fEH0",
            "1zURB1fThVmD1bNY0o6XiDUMsTjLu9IjE4H7sZqF1Ing",
            "1LjPv9qkzUAsfWvoV-3Vifr-xSltbT1ufJPc1iwD_c4k"
        ];

        for (var e = 0; e < extraIds.length; e++) {
            try {
                var openSs = SpreadsheetApp.openById(extraIds[e]);
                if (openSs) spreadSheets.push(openSs);
            } catch (errOpen) { }
        }

        for (var sIdx = 0; sIdx < spreadSheets.length; sIdx++) {
            var ss = spreadSheets[sIdx];
            if (!ss) continue;

            // 1. Kiểm tra sheet THÔNG TIN ĐỊA CHỈ
            var ttdcSheet = ss.getSheetByName("THÔNG TIN ĐỊA CHỈ");
            if (ttdcSheet) {
                var valTtdc = ttdcSheet.getDataRange().getDisplayValues();
                for (var i = 1; i < valTtdc.length; i++) {
                    if (String(valTtdc[i][0]).trim().toUpperCase() === searchCode) {
                        var nameTtdc = String(valTtdc[i][1] || "").trim();
                        var dateTtdc = String(valTtdc[i][3] || "").trim();
                        if (nameTtdc) {
                            return { success: true, name: nameTtdc, date: dateTtdc };
                        }
                    }
                }
            }

            // 2. Kiểm tra CỘT C của sheet 出勤数据 trước tiên (YÊU CẦU TRỰC TIẾP CỦA KHÁCH HÀNG)
            var hrSheet = ss.getSheetByName("出勤数据");
            if (hrSheet) {
                var dataH = hrSheet.getDataRange().getDisplayValues();
                for (var i = 1; i < dataH.length; i++) {
                    var r = dataH[i];
                    var codeInColB = String(r[1] || "").trim().toUpperCase();
                    if (codeInColB === searchCode) {
                        var nameB = String(r[2] || "").trim();
                        var dateB = String(r[3] || "").trim();
                        return { success: true, name: nameB, date: dateB };
                    }
                    for (var c = 0; c < r.length; c++) {
                        var cellVal = String(r[c] || "").trim().toUpperCase();
                        if (cellVal === searchCode) {
                            // LẤY GIÁ TRỊ CỘT C (Index 2)!
                            var nameInColC = String(r[2] || "").trim();
                            var dateInColD = String(r[3] || "").trim();
                            if (nameInColC && nameInColC.toUpperCase() !== searchCode) {
                                return { success: true, name: nameInColC, date: dateInColD };
                            }
                            // Nếu ô trùng chính là Cột C, lấy Cột B (Index 1) hoặc Cột D (Index 3)
                            var altName = String(r[1] || r[3] || r[0] || "").trim();
                            if (altName && altName.toUpperCase() !== searchCode) {
                                return { success: true, name: altName, date: dateInColD };
                            }
                        }
                    }
                }
            }

            // 3. Kiểm tra sheet CHUYÊN CẦN STA5 (Mã Cột A, Tên Cột B)
            var sta5Sheet = ss.getSheetByName("CHUYÊN CẦN STA5");
            if (sta5Sheet) {
                var dataSta = sta5Sheet.getDataRange().getDisplayValues();
                for (var i = 1; i < dataSta.length; i++) {
                    var idSta = String(dataSta[i][0] || "").trim().toUpperCase();
                    if (idSta === searchCode) {
                        var nameSta = String(dataSta[i][1] || "").trim();
                        var dateSta = String(dataSta[i][3] || "").trim();
                        if (nameSta) return { success: true, name: nameSta, date: dateSta };
                    }
                }
            }

            // 4. Kiểm tra sheet NHẬP THẺ / CẤP THẺ
            var nhapTheSheet = ss.getSheetByName("NHẬP THẺ") || ss.getSheetByName("CẤP THẺ");
            if (nhapTheSheet) {
                var dataN = nhapTheSheet.getDataRange().getDisplayValues();
                for (var i = 1; i < dataN.length; i++) {
                    var codeN = String(dataN[i][4] || dataN[i][2] || dataN[i][1] || "").trim().toUpperCase();
                    if (codeN === searchCode) {
                        var nameN = String(dataN[i][5] || dataN[i][3] || dataN[i][2] || "").trim();
                        var dateN = String(dataN[i][6] || dataN[i][5] || "").trim();
                        if (nameN) return { success: true, name: nameN, date: dateN };
                    }
                }
            }
        }

        return { success: false, name: "", date: "" };
    } catch (e) {
        return { success: false, name: "", date: "", message: e.toString() };
    }
}

function saveInTheData(dataArr, deletedCapTheArr, currentUser) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
    } catch (e) {
        return { success: false, message: "Hệ thống đang bận do có người khác đang thao tác. Vui lòng thử lại sau vài giây!" };
    }
    try {
        incrementDataVersion();
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("NHẬP THẺ");
        var capTheSheet = ss.getSheetByName("CẤP THẺ");
        var lsSheet = ss.getSheetByName("LỊCH_SỬ_CẤP_THẺ");
        var lastRow = sheet.getRange("A1:A").getValues().filter(String).length;
        var currentCount = lastRow - 1;
        if (currentCount + dataArr.length > 100) return { success: false, message: "full" };
        sheet.getRange(lastRow + 1, 1, dataArr.length, 12).setValues(dataArr);

        if (capTheSheet) {
            var maThesToClear = dataArr.map(function (r) { return String(r[2]).trim().toUpperCase(); });
            var ctData = capTheSheet.getDataRange().getDisplayValues();
            var historyData = [];
            var timeProcess = new Date().toLocaleString('vi-VN');

            for (var r = ctData.length - 1; r >= 1; r--) {
                var maCT = String(ctData[r][4]).trim().toUpperCase();
                if (maCT && maThesToClear.indexOf(maCT) !== -1) {
                    var rData = ctData[r];
                    var thoiGianYeuCau = timeProcess;
                    historyData.push([
                        thoiGianYeuCau, rData[1], rData[2], rData[3], rData[4], rData[5], rData[6], rData[7], rData[8],
                        timeProcess, currentUser, "Đã In Thẻ", ""
                    ]);
                    capTheSheet.deleteRow(r + 1);
                }
            }
            if (lsSheet && historyData.length > 0) {
                lsSheet.getRange(lsSheet.getLastRow() + 1, 1, historyData.length, historyData[0].length).setValues(historyData);
            }
        }
        SpreadsheetApp.flush();
        var details = dataArr.map(function (r) { return "Ma the: " + r[2] + " - " + r[3]; }).join("\n");
        logAction("Lưu Nhập Thẻ", currentUser, "Đã lưu " + dataArr.length + " dòng nhập thẻ:\n" + details);
        try {
            var now = new Date().toLocaleString('vi-VN');
            var telegramMsg = "✅ [CTQ] LƯU IN THẺ MỚI\n";
            telegramMsg += "Số lượng: " + dataArr.length + " thẻ\n";
            telegramMsg += "-----------------------\n";
            for (var k = 0; k < dataArr.length; k++) {
                var r = dataArr[k];
                telegramMsg += "1: Người Thao Tác: " + (currentUser || "Khách") + "\n";
                telegramMsg += "2: Mã Thẻ: " + r[2] + "\n";
                telegramMsg += "3: Họ Tên : " + r[3] + "\n";
                telegramMsg += "4: Line :\n";
                telegramMsg += "5: Model: " + r[0] + "\n";
                telegramMsg += "6: Bộ Phận: " + r[1] + "\n";
                telegramMsg += "7: Ngày Phát Thẻ: " + r[6] + "\n";
                telegramMsg += "8: Loại Thẻ: " + r[7] + "\n";
                telegramMsg += "14: CTQ 1 : " + r[9] + "\n";
                telegramMsg += "15: CTQ 2 : " + r[10] + "\n";
                telegramMsg += "16: CTQ 3 : " + r[11] + "\n";
                if (k < dataArr.length - 1) telegramMsg += "-----------------------\n";
            }
            sendTelegramNotify(telegramMsg);
        } catch (ez) { }
        return { success: true, newInTheData: sheet.getDataRange().getDisplayValues(), newCapTheData: capTheSheet.getDataRange().getDisplayValues() };
    } catch (e) {
        return { success: false, message: "err" };
    } finally {
        lock.releaseLock();
    }
}

function saveYeuCauCapThe(dataArr, currentUser) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
    } catch (e) {
        return { success: false, message: "Hệ thống đang bận do có người khác đang thao tác. Vui lòng thử lại sau vài giây!" };
    }
    try {
        incrementDataVersion();
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("CẤP THẺ");
        if (!sheet) return { success: false, message: "Không tìm thấy sheet CẤP THẺ!" };

        var existingData = sheet.getDataRange().getDisplayValues();
        var existingMaCodes = {};
        for (var i = 1; i < existingData.length; i++) {
            var existMa = String(existingData[i][4] || "").trim().toUpperCase();
            if (existMa) existingMaCodes[existMa] = true;
        }

        var validRows = [];
        var duplicates = [];
        for (var j = 0; j < dataArr.length; j++) {
            var ma = String(dataArr[j][4] || "").trim().toUpperCase();
            if (existingMaCodes[ma]) {
                duplicates.push(ma);
            } else {
                existingMaCodes[ma] = true;
                validRows.push(dataArr[j]);
            }
        }

        if (duplicates.length > 0 && validRows.length === 0) {
            return { success: false, message: "Tất cả mã thẻ đã tồn tại: " + duplicates.join(", ") };
        }

        if (validRows.length === 0) {
            return { success: false, message: "Không có dữ liệu hợp lệ để lưu!" };
        }

        var lastRow = sheet.getRange("A1:A").getValues().filter(String).length;
        var nextSTT = 1;
        if (lastRow > 1) {
            var lastSTT = parseInt(existingData[existingData.length - 1][0]) || 0;
            nextSTT = lastSTT + 1;
        }

        // Lấy cấu hình duyệt thẻ
        var approvalConfig = getCardApprovalConfig();
        var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

        for (var k = 0; k < validRows.length; k++) {
            validRows[k][0] = nextSTT++;
            // Đảm bảo mảng có đủ 16 phần tử (A-P)
            while (validRows[k].length < 16) validRows[k].push("");
            // Cột K (index 10): Trạng thái = pending_confirm
            validRows[k][10] = "pending_confirm";
            // Cột L (index 11): Người duyệt thẻ (chưa duyệt)
            validRows[k][11] = "";
            // Cột M (index 12): Thời gian duyệt thẻ
            validRows[k][12] = "";
            // Cột N (index 13): Người duyệt in
            validRows[k][13] = "";
            // Cột O (index 14): Thời gian duyệt in
            validRows[k][14] = "";
            // Cột P (index 15): Thời gian tạo yêu cầu + người tạo
            validRows[k][15] = now + " | " + (currentUser || "system");
        }

        sheet.getRange(lastRow + 1, 1, validRows.length, 16).setValues(validRows);
        SpreadsheetApp.flush();

        var details = validRows.map(function(r) { return "Ma the: " + r[4] + " - " + r[5]; }).join("\n");
        logAction("Yêu Cầu Cấp Thẻ", currentUser, "Đã lưu " + validRows.length + " dòng yêu cầu cấp thẻ:\n" + details);

        try {
            var telegramMsg = "🎴 [CTQ] YÊU CẦU CẤP THẺ MỚI\n";
            telegramMsg += "Người Yêu Cầu: " + (currentUser || "Khách") + "\n";
            telegramMsg += "Số lượng: " + validRows.length + " thẻ\n";
            telegramMsg += "-----------------------\n";
            for (var t = 0; t < validRows.length; t++) {
                var row = validRows[t];
                telegramMsg += "STT: " + row[0] + " | Mã: " + row[4] + " | Tên: " + row[5] + " | Model: " + row[2] + " | BP: " + row[1] + "\n";
            }
            sendTelegramNotify(telegramMsg);
        } catch (ez) { }

        var result = { success: true, newCapTheData: sheet.getDataRange().getDisplayValues() };
        if (duplicates.length > 0) {
            result.message = "Đã lưu " + validRows.length + " dòng. " + duplicates.length + " mã trùng bị bỏ qua: " + duplicates.join(", ");
        }
        return result;
    } catch (e) {
        return { success: false, message: "Lỗi hệ thống: " + e.message };
    } finally {
        lock.releaseLock();
    }
}

function deleteYeuCauCapThe(rowIndexes, currentUser) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
    } catch (e) {
        return { success: false, message: "Hệ thống đang bận!" };
    }
    try {
        incrementDataVersion();
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("CẤP THẺ");
        if (!sheet) return { success: false, message: "Không tìm thấy sheet CẤP THẺ!" };

        rowIndexes.sort(function(a, b) { return b - a; });
        var deletedMas = [];

        for (var i = 0; i < rowIndexes.length; i++) {
            var rowIdx = rowIndexes[i] + 1;
            if (rowIdx > 1 && rowIdx <= sheet.getLastRow()) {
                var maThe = sheet.getRange(rowIdx, 5).getDisplayValue();
                deletedMas.push(maThe);
                sheet.deleteRow(rowIdx);
            }
        }

        SpreadsheetApp.flush();
        logAction("Xóa Yêu Cầu Cấp Thẻ", currentUser, "Đã xóa " + deletedMas.length + " dòng: " + deletedMas.join(", "));

        return { success: true, newCapTheData: sheet.getDataRange().getDisplayValues() };
    } catch (e) {
        return { success: false, message: "Lỗi hệ thống: " + e.message };
    } finally {
        lock.releaseLock();
    }
}

// ==================== HỆ THỐNG DUYỆT THẺ NHIỀU CẤP ====================

function getCardApprovalConfig() {
    try {
        var props = PropertiesService.getScriptProperties();
        var configStr = props.getProperty("CARD_APPROVAL_CONFIG");
        if (configStr) {
            var config = JSON.parse(configStr);
            // Backward compatible: nếu config cũ có confirmUsers ở root thì chuyển sang day/night
            if (config.confirmUsers && !config.day) {
                return {
                    day: { confirmUsers: config.confirmUsers, printUsers: config.printUsers },
                    night: { confirmUsers: config.confirmUsers, printUsers: config.printUsers }
                };
            }
            return config;
        }
    } catch (e) { }
    return {
        day: { confirmUsers: [], printUsers: [] },
        night: { confirmUsers: [], printUsers: [] }
    };
}

function saveCardApprovalConfig(config) {
    try {
        incrementDataVersion();
        var props = PropertiesService.getScriptProperties();
        props.setProperty("CARD_APPROVAL_CONFIG", JSON.stringify(config));
        return { success: true };
    } catch (e) {
        return { success: false, message: e.message };
    }
}

function approveCardRequest(rowIndexes, action, currentUser) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
    } catch (e) {
        return { success: false, message: "Hệ thống đang bận!" };
    }
    try {
        incrementDataVersion();
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("CẤP THẺ");
        if (!sheet) return { success: false, message: "Không tìm thấy sheet CẤP THẺ!" };

        var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
        var updated = 0;

        for (var i = 0; i < rowIndexes.length; i++) {
            var sheetRow = rowIndexes[i] + 1; // +1 vì header
            if (sheetRow < 2 || sheetRow > sheet.getLastRow()) continue;

            var currentStatus = sheet.getRange(sheetRow, 11).getDisplayValue(); // Cột K

            if (action === "confirm" && currentStatus === "pending_confirm") {
                // Cấp 1: Xác nhận thẻ
                sheet.getRange(sheetRow, 11).setValue("pending_print");   // K: chuyển trạng thái
                sheet.getRange(sheetRow, 12).setValue(currentUser);       // L: người duyệt
                sheet.getRange(sheetRow, 13).setValue(now);               // M: thời gian duyệt
                updated++;
            } else if (action === "approve_print" && currentStatus === "pending_print") {
                // Cấp 2: Duyệt in thẻ
                sheet.getRange(sheetRow, 11).setValue("approved");        // K: đã duyệt
                sheet.getRange(sheetRow, 14).setValue(currentUser);       // N: người duyệt in
                sheet.getRange(sheetRow, 15).setValue(now);               // O: thời gian duyệt in
                updated++;
            } else if (action === "reject") {
                // Từ chối
                sheet.getRange(sheetRow, 11).setValue("rejected");        // K: từ chối
                if (currentStatus === "pending_confirm") {
                    sheet.getRange(sheetRow, 12).setValue(currentUser);
                    sheet.getRange(sheetRow, 13).setValue(now);
                } else if (currentStatus === "pending_print") {
                    sheet.getRange(sheetRow, 14).setValue(currentUser);
                    sheet.getRange(sheetRow, 15).setValue(now);
                }
                sheet.getRange(sheetRow, 16).setValue("Từ chối bởi " + currentUser + " lúc " + now);
                updated++;
            }
        }

        SpreadsheetApp.flush();

        var actionName = action === "confirm" ? "Xác nhận thẻ" : (action === "approve_print" ? "Duyệt in thẻ" : "Từ chối");
        logAction("Duyệt Cấp Thẻ", currentUser, actionName + " cho " + updated + " dòng");

        return { success: true, updated: updated, newCapTheData: sheet.getDataRange().getDisplayValues() };
    } catch (e) {
        return { success: false, message: "Lỗi: " + e.message };
    } finally {
        lock.releaseLock();
    }
}

function saveNhapTheToThongTinThe(dataArr, currentUser) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
    } catch (e) {
        return { success: false, message: "Hệ thống đang bận do có người khác đang thao tác. Vui lòng thử lại sau vài giây!" };
    }
    try {
        incrementDataVersion();
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var inTheSheet = ss.getSheetByName("NHẬP THẺ");
        var thongTinTheSheet = ss.getSheetByName("THÔNG TIN THẺ");
        if (!thongTinTheSheet) return { success: false, message: "Không tìm thấy sheet THÔNG TIN THẺ" };
        var colEValues = thongTinTheSheet.getRange("E:E").getDisplayValues().flat();
        for (var i = 0; i < dataArr.length; i++) {
            var maThe = dataArr[i][4];
            if (maThe && colEValues.indexOf(String(maThe)) !== -1) return { success: false, message: "Mã thẻ [" + maThe + "] đã tồn tại!" };
        }
        var lastRow = 0;
        for (var r = colEValues.length - 1; r >= 0; r--) { if (colEValues[r] !== "") { lastRow = r + 1; break; } }
        if (lastRow < 2) lastRow = 2;
        thongTinTheSheet.getRange(lastRow + 1, 1, dataArr.length, 18).setValues(dataArr);
        SpreadsheetApp.flush();
        var details = dataArr.map(function (r) { return "Ma the: " + r[4] + " - " + r[5]; }).join("\n");
        logAction("Lưu Nhập Thẻ -> Hệ Thống", currentUser, "Đã chuyển " + dataArr.length + " nhân viên mới vào CSDL THÔNG TIN THẺ:\n" + details);
        try {
            var now = new Date().toLocaleString('vi-VN');
            var telegramMsg = "✅ [CTQ] LƯU NHẬP THẺ MỚI\n";
            telegramMsg += "Số lượng: " + dataArr.length + " thẻ\n";
            telegramMsg += "-----------------------\n";
            for (var k = 0; k < dataArr.length; k++) {
                var r = dataArr[k];
                telegramMsg += "1: Người Thao Tác: " + (currentUser || "Khách") + "\n";
                telegramMsg += "2: Mã Thẻ: " + r[4] + "\n";
                telegramMsg += "3: Họ Tên : " + r[5] + "\n";
                telegramMsg += "4: Line : " + r[0] + "\n";
                telegramMsg += "5: Model: " + r[1] + "\n";
                telegramMsg += "6: Bộ Phận: " + r[2] + "\n";
                telegramMsg += "7: Ngày Phát Thẻ: " + r[7] + "\n";
                telegramMsg += "8: Loại Thẻ: " + r[8] + "\n";
                telegramMsg += "9: Cấp Lần 1: " + r[10] + "\n";
                telegramMsg += "10: Cấp Lần 2: " + r[11] + "\n";
                telegramMsg += "11: Lý Do Cấp L2: " + r[12] + "\n";
                telegramMsg += "12: Cấp Lần 3: " + r[13] + "\n";
                telegramMsg += "13: Lý Do Cấp L3: " + r[14] + "\n";
                telegramMsg += "14: CTQ 1 : " + r[15] + "\n";
                telegramMsg += "15: CTQ 2 : " + r[16] + "\n";
                telegramMsg += "16: CTQ 3 : " + r[17] + "\n";
                if (k < dataArr.length - 1) telegramMsg += "-----------------------\n";
            }
            sendTelegramNotify(telegramMsg);
        } catch (ez) { }
        return { success: true, newInTheData: inTheSheet.getDataRange().getDisplayValues(), newThongTinThe: thongTinTheSheet.getDataRange().getDisplayValues() };
    } catch (e) {
        return { success: false, message: e.toString() };
    } finally {
        lock.releaseLock();
    }
}

function timKiemChinhSua(maThe) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("THÔNG TIN THẺ");
    if (!sheet) return { success: false, message: "no_sheet" };
    var finder = sheet.getRange("E:E").createTextFinder(maThe.trim()).matchEntireCell(true).findNext();
    if (finder) {
        var row = finder.getRow();
        return { success: true, row: row, data: sheet.getRange(row, 1, 1, 18).getDisplayValues()[0] };
    }
    return { success: false, message: "not_found" };
}
function luuChinhSuaThongTin(row, dataArr, currentUser) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
    } catch (e) {
        throw new Error("Hệ thống đang bận do có người khác đang thao tác. Vui lòng thử lại sau vài giây!");
    }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("THÔNG TIN THẺ");
        sheet.getRange(row, 1, 1, 18).setValues([dataArr]);
        SpreadsheetApp.flush();
        logAction("Lưu Chỉnh Sửa Thông Tin", currentUser, "Dòng: " + row + " | Mã thẻ: " + dataArr[4] + " | Họ tên: " + dataArr[5]);
        return sheet.getDataRange().getDisplayValues();
    } finally {
        lock.releaseLock();
    }
}

function xoaChinhSuaThongTin(row, currentUser) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
    } catch (e) {
        throw new Error("Hệ thống đang bận do có người khác đang thao tác. Vui lòng thử lại sau vài giây!");
    }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("THÔNG TIN THẺ");
        var data = sheet.getDataRange().getDisplayValues();
        var rIdx = row - 1;
        var detail = "Dòng: " + row;
        if (data[rIdx]) {
            detail += " | Mã Thẻ: " + data[rIdx][4] + " - Tên: " + data[rIdx][5];
        }
        sheet.deleteRow(row);
        SpreadsheetApp.flush();
        logAction("Xoá Dữ Liệu Nhân Viên", currentUser, "Đã xoá nhân viên:\n" + detail);
        return sheet.getDataRange().getDisplayValues();
    } finally {
        lock.releaseLock();
    }
}

function xoaDanhSachThongTinTheByMaThe(maTheArr, currentUser) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(15000);
    } catch (e) {
        throw new Error("Hệ thống đang bận. Vui lòng thử lại sau!");
    }
    try {
        if (!maTheArr || !maTheArr.length) return { success: false, message: "Chưa chọn mã thẻ nào!" };
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("THÔNG TIN THẺ");
        if (!sheet) return { success: false, message: "Không tìm thấy sheet THÔNG TIN THẺ" };

        var data = sheet.getDataRange().getDisplayValues();
        var targetSet = {};
        maTheArr.forEach(function (m) { targetSet[String(m).trim().toUpperCase()] = true; });

        var rowsToDelete = [];
        for (var i = 1; i < data.length; i++) {
            var ma = String(data[i][4] || '').trim().toUpperCase();
            if (targetSet[ma]) {
                rowsToDelete.push(i + 1); // 1-indexed row number in Sheet
            }
        }

        if (rowsToDelete.length === 0) return { success: false, message: "Không tìm thấy mã thẻ tương ứng trong Sheet!" };

        // Xóa từ dưới lên trên để không làm lệch chỉ số dòng
        rowsToDelete.sort(function (a, b) { return b - a; });
        rowsToDelete.forEach(function (r) {
            sheet.deleteRow(r);
        });

        SpreadsheetApp.flush();
        logAction("Xoá Danh Sách Thẻ", currentUser, "Đã xoá " + rowsToDelete.length + " nhân viên trong THÔNG TIN THẺ: " + maTheArr.join(", "));
        return { success: true, count: rowsToDelete.length, newData: sheet.getDataRange().getDisplayValues() };
    } finally {
        lock.releaseLock();
    }
}
// HÀM KIỂM TRA QUYỀN REAL-TIME TỪ TRÌNH DUYỆT
function checkUserRealTime(username) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var props = PropertiesService.getScriptProperties();

    // KIỂM TRA XEM USER CÓ BỊ KICKED KHÔNG
    var kickedList = props.getProperty("KICKED_USERS");
    if (kickedList) {
        var kicked = kickedList.split(",");
        if (kicked.indexOf(username) !== -1) {
            return { exists: false, kicked: true };
        }
    }

    // LƯU Ý QUAN TRỌNG: Thay chữ "USERS" bằng đúng tên Trang tính (Sheet) chứa tài khoản đăng nhập của bạn nhé
    var sheet = ss.getSheetByName("USERS");

    if (!sheet) return { exists: false };

    var data = sheet.getDataRange().getValues();

    // Chạy vòng lặp tìm tên đăng nhập
    for (var i = 1; i < data.length; i++) {
        // Giả sử Cột A (index 0) là Username, Cột C (index 2) là Role, Cột D (index 3) là Programs
        if (data[i][0] == username) {
            return {
                exists: true,
                role: data[i][2],
                perms: data[i][3],
                dataVersion: parseInt(props.getProperty('GLOBAL_DATA_VERSION') || '0')
            };
        }
    }

    // Nếu quét hết bảng mà không thấy tên user này -> Trả về false để kích văng
    return { exists: false };
}

// HÀM KICK USER LOGOUT NGAY LẬP TỨC
function kickUserOut(username) {
    try {
        var props = PropertiesService.getScriptProperties();
        var kickedList = props.getProperty("KICKED_USERS") || "";

        // THÊM USERNAME VÀO DANH SÁCH KICKED
        if (kickedList === "") {
            kickedList = username;
        } else if (kickedList.indexOf(username) === -1) {
            kickedList += "," + username;
        }

        props.setProperty("KICKED_USERS", kickedList);
        return { success: true, message: "User " + username + " has been kicked out" };
    } catch (e) {
        return { success: false, message: "Error: " + e.toString() };
    }
}

// HÀM CLEAR SESSION KHI USER LOGOUT BÌNH THƯỜNG
function clearKickedStatus(username) {
    try {
        var props = PropertiesService.getScriptProperties();
        var kickedList = props.getProperty("KICKED_USERS") || "";

        // LOẠI BỎ USERNAME KHỎI DANH SÁCH KICKED
        if (kickedList !== "") {
            var kicked = kickedList.split(",");
            var idx = kicked.indexOf(username);
            if (idx !== -1) {
                kicked.splice(idx, 1);
                props.setProperty("KICKED_USERS", kicked.join(","));
            }
        }
        return { success: true };
    } catch (e) {
        return { success: false };
    }
}

// HÀM RECORD USER ĐĂNG NHẬP
function recordUserLogin(username) {
    try {
        var props = PropertiesService.getScriptProperties();
        var loginList = props.getProperty("LOGGED_IN_USERS") || "";
        var loginData = props.getProperty("USER_LOGIN_" + username) || "";

        // THẾ ĐĂNG NHẬP RỒI THÌ CẬP NHẬT THỜI GIAN
        var now = new Date().toLocaleString('vi-VN');
        var nowMs = new Date().getTime().toString();
        if (loginList.indexOf(username) !== -1) {
            props.setProperty("USER_LOGIN_" + username, now);
            props.setProperty("LAST_SEEN_" + username, nowMs);
            return { success: true };
        }

        // NẾU CHƯA THÌ THÊM MỚI
        if (loginList === "") {
            loginList = username;
        } else {
            loginList += "," + username;
        }

        props.setProperty("LOGGED_IN_USERS", loginList);
        props.setProperty("USER_LOGIN_" + username, now);
        props.setProperty("LAST_SEEN_" + username, nowMs);
        logAction("Đăng Nhập", username, "Đã truy cập vào hệ thống");
        try {
            var now = new Date().toLocaleString('vi-VN');
            sendTelegramNotify("[CTQ] Đăng Nhập Hệ Thống\nTài khoản: " + username + "\nThời gian: " + now);
        } catch (ez) { }
        return { success: true };
    } catch (e) {
        return { success: false };
    }
}

// HÀM XÓA USER KHỎI DANH SÁCH ĐĂNG NHẬP
function recordUserLogout(username) {
    try {
        var props = PropertiesService.getScriptProperties();
        var loginList = props.getProperty("LOGGED_IN_USERS") || "";

        if (loginList !== "") {
            var users = loginList.split(",");
            var idx = users.indexOf(username);
            if (idx !== -1) {
                users.splice(idx, 1);
                props.setProperty("LOGGED_IN_USERS", users.join(","));
            }
        }

        props.deleteProperty("USER_LOGIN_" + username);
        props.deleteProperty("LAST_SEEN_" + username);
        try {
            var now = new Date().toLocaleString('vi-VN');
            sendTelegramNotify("[CTQ] Đăng Xuất Hệ Thống\nTài khoản: " + username + "\nThời gian: " + now);
        } catch (ez) { }
        return { success: true };
    } catch (e) {
        return { success: false };
    }
}

// 2. HÀM LẤY DANH SÁCH USER (MÁY CHÉM BÓNG MA BẢN NÂNG CẤP TỐI THƯỢNG)
function getLoggedInUsers() {
    try {
        var props = PropertiesService.getScriptProperties();
        var loginList = props.getProperty("LOGGED_IN_USERS") || "";
        var result = [];

        if (loginList !== "") {
            var users = loginList.split(",");
            var validUsers = [];
            var nowTime = new Date().getTime(); // Giờ hiện tại

            for (var i = 0; i < users.length; i++) {
                var user = users[i];
                if (!user) continue;

                var lastSeen = props.getProperty("LAST_SEEN_" + user);

                // VÁ LỖ HỔNG: Nếu KHÔNG CÓ nhịp tim (!lastSeen) HOẶC nhịp tim quá cũ (> 2 phút) -> CHÉM!
                if (!lastSeen || (nowTime - parseInt(lastSeen) > 1000)) {
                    props.deleteProperty("USER_LOGIN_" + user);
                    props.deleteProperty("LAST_SEEN_" + user);
                } else {
                    // Nếu vẫn sống thì cho vào danh sách hiển thị
                    validUsers.push(user);
                    var loginTime = props.getProperty("USER_LOGIN_" + user);
                    result.push({
                        username: user,
                        loginTime: loginTime || "Unknown"
                    });
                }
            }

            // CHỐT SỔ: Lưu lại danh sách mới
            if (validUsers.length !== users.length) {
                props.setProperty("LOGGED_IN_USERS", validUsers.join(","));
            }
        }
        return { success: true, data: result };
    } catch (e) {
        return { success: false, data: [] };
    }
}

// HÀM GỬI THÔNG BÁO CHO USER
function sendNotificationToUser(targetUser, message) {
    try {
        var props = PropertiesService.getScriptProperties();
        var timestamp = new Date().toLocaleString('vi-VN');
        var notificationData = JSON.stringify({ message: message, time: timestamp });

        props.setProperty("NOTIFICATION_" + targetUser, notificationData);
        return { success: true, message: "Đã gửi thông báo cho " + targetUser };
    } catch (e) {
        return { success: false, message: "Lỗi: " + e.toString() };
    }
}

// HÀM LẤY THÔNG BÁO CHO USER
function getNotificationForUser(username) {
    try {
        var props = PropertiesService.getScriptProperties();
        var notification = props.getProperty("NOTIFICATION_" + username);

        if (notification) {
            return { success: true, hasNotification: true, data: JSON.parse(notification) };
        }
        return { success: true, hasNotification: false, data: null };
    } catch (e) {
        return { success: false, hasNotification: false };
    }
}

// HÀM XÓA THÔNG BÁO SAU KHI ĐÃ ĐỌC
function clearNotificationForUser(username) {
    try {
        var props = PropertiesService.getScriptProperties();
        props.deleteProperty("NOTIFICATION_" + username);
        return { success: true };
    } catch (e) {
        return { success: false };
    }
}




// ====== FORM XÁC NHẬN LỖI - LẤY DỮ LIỆU TỪ SHEET DATA (TỐI ƯU) ======
function getXacNhanLoiData(fromDate, toDate, line, model, ca, status) {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("DATA");
        if (!sheet) return { success: false, message: "Không tìm thấy sheet DATA" };

        // ✅ TỐI ƯU: Lấy từ hàng 3 trở đi (index 2), bỏ header
        var lastRow = sheet.getLastRow();
        if (lastRow < 3) return { success: true, data: [], stats: { totalErrors: 0, confirmed: 0, notConfirmed: 0, rejected: 0, agreementRate: "0%" } };

        // ✅ TỐI ƯU: Lấy dữ liệu từ cột A đến S (19 cột)
        var data = sheet.getRange(3, 1, lastRow - 2, 19).getDisplayValues();
        var result = [];
        var stats = {
            totalErrors: 0,
            confirmed: 0,
            notConfirmed: 0,
            rejected: 0,
            agreementRate: "0%",
            errorCount: 0
        };

        // DÒ DỮ LIỆU VÀ LỌC
        for (var i = 0; i < data.length; i++) {
            var row = data[i];

            // Cột: TÊN TRẠM(0), MÃ THẺ(1), HỌ TÊN(2), TIME(3), LẦN TEST(4), LỖI(5), MÃ LỖI(6), TÊN LỖI(7), THỰC HIỆN(8), LINE(9), XÁC NHẬN(10), NGÀY(11), CA(12), MODEL(13), Status(14) ... CHO XÁC NHẬN(18)
            var maLoi = String(row[6] || "").trim();

            // ✅ BỎ QUA CTQ-00
            if (maLoi === "CTQ-00" || maLoi === "") continue;

            var rowDate = row[11] || "";
            var rowLine = row[9] || "";
            var rowModel = row[13] || "";
            var rowCa = row[12] || "";
            var rowStatus = row[14] || ""; // THỰC HIỆN = Status (Lọc ở cột O)
            var facaText = row[16] || ""; // Cột Q: FACA
            var userInfo = row[17] || ""; // Cột R: User Info (admin | time | ms)
            var finalWaitTime = row[18] || ""; // Cột S: Final Wait Time

            // LỌC THEO NGÀY
            if (fromDate && toDate) {
                try {
                    var rDate;
                    var parts = String(rowDate).split('/');
                    if (parts.length === 3) {
                        rDate = new Date(parts[2], parts[1] - 1, parts[0]);
                    } else {
                        rDate = new Date(rowDate);
                    }
                    var fDate = new Date(fromDate);
                    var tDate = new Date(toDate);

                    fDate.setHours(0, 0, 0, 0);
                    tDate.setHours(23, 59, 59, 999);

                    if (isNaN(rDate.getTime()) || rDate < fDate || rDate > tDate) continue;
                } catch (e) { continue; }
            }

            // LỌC THEO LINE
            if (line && line !== "ALL" && rowLine !== line) continue;

            // LỌC THEO MODEL
            if (model && model !== "ALL" && rowModel !== model) continue;

            // LỌC THEO CA
            if (ca && ca !== "ALL" && rowCa !== ca) continue;

            // LỌC THEO STATUS
            if (status && status !== "ALL") {
                if (status === "confirmed" && !rowStatus.includes("Đã Xác Nhận")) continue;
                if (status === "notConfirmed" && !rowStatus.includes("Chưa Xác Nhận")) continue;
                if (status === "rejected" && !rowStatus.includes("Từ Chối")) continue;
            }

            var choXacNhanStr = "";
            if (rowStatus.includes("Đã Xác Nhận") || rowStatus.includes("Từ Chối")) {
                choXacNhanStr = finalWaitTime;
            } else {
                if (userInfo) {
                    var parts = userInfo.split("|");
                    if (parts.length >= 3) {
                        var ts = parseInt(parts[2].trim());
                        if (!isNaN(ts)) {
                            var diffMs = new Date().getTime() - ts;
                            var diffMins = Math.floor(diffMs / 60000);
                            if (diffMins < 60) {
                                choXacNhanStr = diffMins + " phút";
                            } else {
                                var diffHours = Math.floor(diffMins / 60);
                                var remainMins = diffMins % 60;
                                choXacNhanStr = diffHours + " giờ " + remainMins + " phút";
                            }
                        }
                    }
                }
            }

            // THÊM VÀO KẾT QUẢ
            result.push({
                rowIndex: i + 3,
                stt: result.length + 1,
                tramName: row[0],
                maThe: row[1],
                hoTen: row[2],
                lanTest: row[4],
                loi: row[5],
                maLoi: maLoi,
                tenLoi: row[7],
                thucHien: row[8],
                line: row[9],
                xacNhan: row[10],
                ngay: row[11],
                ca: row[12],
                model: row[13],
                trangThai: rowStatus,
                facaText: facaText,
                choXacNhan: choXacNhanStr
            });

            // CẬP NHẬT STATS
            stats.totalErrors++;
            if (rowStatus.includes("Đã Xác Nhận")) stats.confirmed++;
            else if (rowStatus.includes("Chưa Xác Nhận") || (!rowStatus.includes("Đã Xác Nhận") && !rowStatus.includes("Từ Chối"))) stats.notConfirmed++;
            else if (rowStatus.includes("Từ Chối")) stats.rejected++;
        }

        // TÍNH %
        if (stats.totalErrors > 0) {
            stats.agreementRate = Math.round((stats.confirmed / stats.totalErrors) * 100) + "%";
        }

        return {
            success: true,
            data: result,
            stats: stats
        };
    } catch (e) {
        return { success: false, message: "Lỗi: " + e.toString() };
    }
}

function xacNhanLoiCoTaiKhoan(rowIndex, statusValue, username, password, facaText, imgBase64, imgName, imgMime) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
    } catch (e) {
        return { success: false, message: "Hệ thống đang bận do có người khác đang thao tác. Vui lòng thử lại sau vài giây!" };
    }
    try {
        incrementDataVersion();
        var sheetUsers = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USERS");
        var dataUsers = sheetUsers.getDataRange().getDisplayValues();
        var isValid = false;
        var userNameDisplay = username;
        for (var i = 1; i < dataUsers.length; i++) {
            if (String(dataUsers[i][0]).trim() === String(username).trim() && String(dataUsers[i][1]).trim() === String(password).trim()) {
                isValid = true;

                break;
            }
        }
        if (!isValid) {
            return { success: false, message: "Tên đăng nhập hoặc mật khẩu không đúng!" };
        }

        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DATA");
        var finalFacaText = facaText;

        if (imgBase64 && imgBase64 !== "") {
            var maThe = sheet.getRange(rowIndex, 2).getValue() || "Unknown";
            var line = sheet.getRange(rowIndex, 10).getValue() || "Unknown";
            var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy");

            var ext = ".png";
            if (imgName && imgName.indexOf(".") > -1) {
                ext = imgName.substring(imgName.lastIndexOf("."));
            }
            var newImgName = userNameDisplay + "_" + maThe + "_" + dateStr + "_" + line + ext;

            var blob = Utilities.newBlob(Utilities.base64Decode(imgBase64), imgMime || "image/png", newImgName);
            var folder = DriveApp.getFolderById("1xwuCx-8K1qPQDyyo8-N-jTohFkjAogcM");
            var file = folder.createFile(blob);

            try {
                file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            } catch (errSharing) {
                // Ignore sharing error if workspace restricts public sharing
            }

            var fileUrl = file.getUrl();
            finalFacaText += "\n\nLink Ảnh Đào Tạo: " + fileUrl;
        }

        var displayStatus = userNameDisplay + " Đã Xác Nhận";

        // Cột O(15): Trạng Thái
        sheet.getRange(rowIndex, 15).setValue(displayStatus);

        // Cột Q(17): FACA
        sheet.getRange(rowIndex, 17).setValue(finalFacaText);

        // Tính thời gian chốt Cột S(19)
        var userInfo = sheet.getRange(rowIndex, 18).getValue();
        if (userInfo) {
            var parts = String(userInfo).split("|");
            if (parts.length >= 3) {
                var ts = parseInt(parts[2].trim());
                if (!isNaN(ts)) {
                    var diffMs = new Date().getTime() - ts;
                    var diffMins = Math.floor(diffMs / 60000);
                    var finalWaitTime = "";
                    if (diffMins < 60) {
                        finalWaitTime = diffMins + " phút";
                    } else {
                        var diffHours = Math.floor(diffMins / 60);
                        var remainMins = diffMins % 60;
                        finalWaitTime = diffHours + " giờ " + remainMins + " phút";
                    }
                    sheet.getRange(rowIndex, 19).setValue(finalWaitTime);
                }
            }
        }

        var maTheLog = sheet.getRange(rowIndex, 2).getValue() || "N/A";
        var lineLog = sheet.getRange(rowIndex, 10).getValue() || "N/A";
        var tenLoiLog = sheet.getRange(rowIndex, 8).getValue() || "N/A";
        var chiTiet = "Line: " + lineLog + "\nMã thẻ: " + maTheLog + "\nLỗi: " + tenLoiLog + "\nFACA: " + finalFacaText;
        logAction("✅ XÁC NHẬN LỖI", userNameDisplay, chiTiet);

        return { success: true };
    } catch (e) {
        return { success: false, message: "Lỗi hệ thống (Có thể do chưa cấp quyền Google Drive): " + e.toString() };
    } finally {
        lock.releaseLock();
    }
}


function xacNhanLoiTrenSheet(rowIndex, statusValue) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
    } catch (e) {
        return { success: false, message: "Hệ thống đang bận do có người khác đang thao tác. Vui lòng thử lại sau vài giây!" };
    }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DATA");
        // Cột O là cột 15, Cột R là 18, Cột S là 19
        sheet.getRange(rowIndex, 15).setValue(statusValue || "Đã Xác Nhận");

        // Tính thời gian chốt (giờ cuối cùng)
        var userInfo = sheet.getRange(rowIndex, 18).getValue();
        if (userInfo) {
            var parts = String(userInfo).split("|");
            if (parts.length >= 3) {
                var ts = parseInt(parts[2].trim());
                if (!isNaN(ts)) {
                    var diffMs = new Date().getTime() - ts;
                    var diffMins = Math.floor(diffMs / 60000);
                    var finalWaitTime = "";
                    if (diffMins < 60) {
                        finalWaitTime = diffMins + " phút";
                    } else {
                        var diffHours = Math.floor(diffMins / 60);
                        var remainMins = diffMins % 60;
                        finalWaitTime = diffHours + " giờ " + remainMins + " phút";
                    }
                    sheet.getRange(rowIndex, 19).setValue(finalWaitTime);
                }
            }
        }

        return { success: true };
    } catch (e) {
        return { success: false, message: e.toString() };
    } finally {
        lock.releaseLock();
    }
}

// HÀM LẤY CÁC OPTION DUYÊN (LINE, MODEL, CA) TỪ SHEET DATA
function getXacNhanLoiFilters() {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("DATA");
        if (!sheet) return { success: false };

        var data = sheet.getDataRange().getDisplayValues();
        var lines = new Set();
        var models = new Set();
        var cas = new Set(['Day', 'Night']); // Mặc định

        for (var i = 1; i < data.length; i++) {
            var lineVal = data[i][9] ? String(data[i][9]).trim() : "";
            var modelVal = data[i][13] ? String(data[i][13]).trim() : "";
            var caVal = data[i][12] ? String(data[i][12]).trim() : "";

            if (lineVal && !['LINE', 'Line', 'Line 线别'].includes(lineVal)) lines.add(lineVal);
            if (modelVal && !['MODEL', 'Model', 'Model 机型'].includes(modelVal)) models.add(modelVal);
            if (caVal && !['CA', 'Ca'].includes(caVal)) cas.add(caVal);
        }

        return {
            success: true,
            lines: Array.from(lines).sort(),
            models: Array.from(models).sort(),
            cas: Array.from(cas).sort()
        };
    } catch (e) {
        return { success: false };
    }
}

// HÀM XUẤT EXCEL - XÁC NHẬN LỖI
function exportXacNhanLoiToExcel(fromDate, toDate, line, model, ca, status) {
    try {
        var result = getXacNhanLoiData(fromDate, toDate, line, model, ca, status);
        if (!result.success) return { success: false };

        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var exportSheet = ss.insertSheet("EXPORT_XacNhanLoi_" + new Date().getTime());

        // HEADER
        var headers = ["STT", "TRẠM", "MÃ THẺ", "HỌ TÊN", "MÃ LỖI", "TÊN LỖI", "LINE", "NGÀY", "CA", "QUẢN LÝ XÁC NHẬN", "TRẠNG THÁI", "CÁI TIỀN", "CHỈ XÁC NHẬN"];
        exportSheet.appendRow(headers);

        // DỮ LIỆU
        for (var i = 0; i < result.data.length; i++) {
            var row = result.data[i];
            exportSheet.appendRow([
                row.stt,
                row.tramName,
                row.maThe,
                row.hoTen,
                row.maLoi,
                row.tenLoi,
                row.line,
                row.ngay,
                row.ca,
                row.thucHien,
                row.xacNhan,
                "", // CẢI TIẾN
                ""  // CHỈ XÁC NHẬN
            ]);
        }

        // FOOTER STATS
        exportSheet.appendRow([]);
        exportSheet.appendRow(["THỐNG KÊ:", "Tổng Lỗi: " + result.stats.totalErrors, "Đã Xác Nhận: " + result.stats.confirmed, "Chưa Xác Nhận: " + result.stats.notConfirmed, "Từ Chối: " + result.stats.rejected, "Tỷ Lệ Đồng Ý: " + result.stats.agreementRate]);

        // FORMAT
        exportSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#0f5132").setFontColor("white");
        exportSheet.autoResizeColumns(1, headers.length);

        return { success: true, sheetName: exportSheet.getName() };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

// ====== TỐI ƯU: GỘP TẤT CẢ SYSTEM STATUS CHECKS THÀNH 1 REQUEST ======
function getSystemStatus(username, sessionId) {
    try {
        if (!username) return { success: false, status: "invalid" };

        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var props = PropertiesService.getScriptProperties();

        // 0. KIỂM TRA SESSION ID
        if (sessionId) {
            var serverSessionId = props.getProperty("SESSION_ID_" + username);
            if (serverSessionId && serverSessionId !== sessionId) {
                return { success: true, kicked: true, status: "session_expired" };
            }
        }

        // 1. KIỂM TRA BỊ KICKED HOẶC LOCKED
        var isKicked = false;
        var kickedList = props.getProperty("KICKED_USERS") || "";
        if (kickedList && kickedList.indexOf(username) !== -1) {
            isKicked = true;
        }

        var lockedStr = props.getProperty("LOCKED_USERS");
        if (lockedStr) {
            try {
                var locks = JSON.parse(lockedStr);
                if (locks[username]) {
                    if (locks[username].until > new Date().getTime()) {
                        return { success: true, kicked: true, status: "locked", lockInfo: locks[username] };
                    } else {
                        delete locks[username];
                        props.setProperty("LOCKED_USERS", JSON.stringify(locks));
                    }
                }
            } catch (e) { }
        }

        // 2. KIỂM TRA QUYỀN & ROLE
        var sheet = ss.getSheetByName("USERS");
        var userData = { exists: false, role: "", perms: "" };
        if (sheet) {
            var data = sheet.getDataRange().getValues();
            for (var i = 1; i < data.length; i++) {
                if (data[i][0] == username) {
                    userData = { exists: true, role: data[i][2], perms: data[i][3] };
                    break;
                }
            }
        }

        // 3. LẤY THÔNG BÁO
        var notification = props.getProperty("NOTIFICATION_" + username);
        var hasNotification = false;
        var notifData = null;
        if (notification) {
            hasNotification = true;
            notifData = JSON.parse(notification);
        }

        var usernameUpper = String(username).trim().toUpperCase();
        var hasUnreadLienHe = props.getProperty("UNREAD_LIENHE_" + usernameUpper) === "true";

        // CẬP NHẬT NHỊP TIM
        var nowTime = new Date().getTime();
        props.setProperty("LAST_SEEN_" + username, nowTime.toString());

        // ĐẢM BẢO BẢN THÂN LUÔN CÓ TRONG DANH SÁCH (TỰ PHỤC HỒI NẾU BỊ MẤT)
        var loginListStr = props.getProperty("LOGGED_IN_USERS") || "";
        var currentUsers = loginListStr ? loginListStr.split(",") : [];
        if (currentUsers.indexOf(username) === -1) {
            currentUsers.push(username);
            // Nếu mất thời gian đăng nhập thì gán lại
            if (!props.getProperty("USER_LOGIN_" + username)) {
                props.setProperty("USER_LOGIN_" + username, new Date().toLocaleString('vi-VN'));
            }
        }

        // 4. LẤY DANH SÁCH USER ONLINE VÀ DỌN DẸP BÓNG MA
        var loggedInUsers = [];
        var validUsers = [];

        for (var j = 0; j < currentUsers.length; j++) {
            var u = currentUsers[j];
            if (!u) continue;

            var lastSeen = props.getProperty("LAST_SEEN_" + u);
            // Hạn chót là 15 giây (refresh mỗi 5s)
            if (!lastSeen || (nowTime - parseInt(lastSeen) > 15000)) {
                props.deleteProperty("USER_LOGIN_" + u);
                props.deleteProperty("LAST_SEEN_" + u);
            } else {
                if (validUsers.indexOf(u) === -1) {
                    validUsers.push(u);
                    var loginTime = props.getProperty("USER_LOGIN_" + u);
                    loggedInUsers.push({
                        username: u,
                        loginTime: loginTime || "Unknown"
                    });
                }
            }
        }

        // CẬP NHẬT LẠI LIST NẾU CÓ SỰ THAY ĐỔI
        var newLoginListStr = validUsers.join(',');
        if (newLoginListStr !== loginListStr) {
            props.setProperty('LOGGED_IN_USERS', newLoginListStr);
        }

        var dataVersion = parseInt(props.getProperty('GLOBAL_DATA_VERSION') || '0');
        var chatVersion = parseInt(props.getProperty('GLOBAL_CHAT_VERSION') || '0');

        var docProps = PropertiesService.getDocumentProperties();

        return {
            success: true,
            kicked: isKicked,
            user: userData,
            hasNotification: hasNotification,
            notification: notifData,
            hasUnreadLienHe: hasUnreadLienHe,
            loggedInUsers: loggedInUsers,
            timestamp: new Date().getTime(),
            dataVersion: dataVersion,
            chatVersion: chatVersion,
            marqueeText: docProps.getProperty('MARQUEE_TEXT') || '',
            marqueeExpiry: docProps.getProperty('MARQUEE_EXPIRY') || '',
            marqueeSpeed: docProps.getProperty('MARQUEE_SPEED') || '20',
            marqueeColor: docProps.getProperty('MARQUEE_COLOR') || '#ff0000',
            marqueeSize: docProps.getProperty('MARQUEE_SIZE') || '16'
        };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function changeUserPassword(username, oldPass, newPass) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
    } catch (e) {
        return { success: false, message: "Hệ thống đang bận do có người khác đang thao tác. Vui lòng thử lại sau vài giây!" };
    }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('USERS');
        var data = sheet.getDataRange().getDisplayValues();

        for (var i = 1; i < data.length; i++) {
            // Tìm đúng dòng của User
            if (data[i][0] === username) {
                // Kiểm tra mật khẩu cũ (cột index 1)
                if (data[i][1] === oldPass) {
                    sheet.getRange(i + 1, 2).setValue(newPass);
                    logAction("Đổi Mật Khẩu", username, "Đã đổi mật khẩu thành công.");
                    return { success: true };
                } else {
                    return { success: false, message: "Mật khẩu cũ không chính xác!" };
                }
            }
        }
        return { success: false, message: "Không tìm thấy tài khoản!" };
    } finally {
        lock.releaseLock();
    }
}

function recoverUserPassword(username, department) {
    try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USERS");
        if (!sheet) return { success: false, message: "Lỗi hệ thống: Không tìm thấy sheet USERS!" };
        var data = sheet.getDataRange().getDisplayValues();
        var user = username.toString().trim();
        var dept = department.toString().trim().toUpperCase();
        for (var i = 1; i < data.length; i++) {
            if (data[i][0].toString().trim() === user) {
                var sheetDept = data[i][4] ? data[i][4].toString().trim().toUpperCase() : "";
                if (sheetDept === dept) {
                    return { success: true, password: data[i][1] };
                } else {
                    return { success: false, message: "Tên bộ phận không chính xác!" };
                }
            }
        }
        return { success: false, message: "Tài khoản không tồn tại!" };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

// ==========================================
// CÁC HÀM XỬ LÝ QUỸ LỢN (NỘP PHẠT)
// ==========================================
function saveTienLon(dataRow, currentUser) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
    } catch (e) {
        throw new Error("Hệ thống đang bận do có người khác đang thao tác. Vui lòng thử lại sau vài giây!");
    }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('QUY_LON');
        sheet.appendRow(dataRow);
        logAction("Ghi Nhận Quỹ Lợn", currentUser, "NV: " + dataRow[1] + " | Lý do: " + dataRow[2] + " | Đã nộp: " + dataRow[3] + " | Nợ: " + dataRow[4]);
        return sheet.getDataRange().getDisplayValues();
    } finally {
        lock.releaseLock();
    }
}

function deleteTienLon(rowIndex, currentUser) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
    } catch (e) {
        throw new Error("Hệ thống đang bận do có người khác đang thao tác. Vui lòng thử lại sau vài giây!");
    }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("QUY_LON");
        var data = sheet.getDataRange().getDisplayValues();
        var rIdx = rowIndex - 1;
        var detail = "Dòng: " + rowIndex;
        if (data[rIdx]) {
            detail += " | NV: " + data[rIdx][1] + " | Lý do: " + data[rIdx][3] + " | Đã nộp: " + data[rIdx][4] + " | Nợ: " + data[rIdx][5];
        }
        sheet.deleteRow(rowIndex);
        logAction("Xoá Giao Dịch Lợn", currentUser, "Đã xoá:\\n" + detail);
        return sheet.getDataRange().getDisplayValues();
    } finally {
        lock.releaseLock();
    }
}

// HÀM ĐỒNG BỘ SIÊU TỐC CẢ QUỸ LỢN VÀ QUYỀN USER (CẬP NHẬT 24/24)
function syncRealtimeAll(currentUser) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var quyLon = ss.getSheetByName("QUY_LON").getDataRange().getDisplayValues();
    var users = ss.getSheetByName("USERS").getDataRange().getDisplayValues();

    var perms = "";
    // Quét tìm quyền mới nhất của người đang đăng nhập
    for (var i = 1; i < users.length; i++) {
        if (users[i][0] === currentUser) {
            perms = users[i][3] || "";
            break;
        }
    }

    return { quyLon: quyLon, perms: perms, users: users };
}
// HÀM MỔ LỢN (XÓA SẠCH QUỸ)
function moConLon(currentUser) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
    } catch (e) {
        throw new Error("Hệ thống đang bận do có người khác đang thao tác. Vui lòng thử lại sau vài giây!");
    }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("QUY_LON");
        var lastRow = sheet.getLastRow();
        if (lastRow > 1) {
            sheet.deleteRows(2, lastRow - 1);
        }
        logAction("MỔ LỢN (LÀM SẠCH QUỸ)", currentUser, "Đã reset toàn bộ dữ liệu quỹ lợn.");
        return sheet.getDataRange().getDisplayValues();
    } finally {
        lock.releaseLock();
    }
}
// HÀM TRẢ NỢ THÔNG MINH: TỰ ĐỘNG TRỪ LÙI & GHI NHẬN NGƯỜI THU
function processTraNoThongMinh(maThe, payAmount, nguoiThu) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
    } catch (e) {
        throw new Error("Hệ thống đang bận do có người khác đang thao tác. Vui lòng thử lại sau vài giây!");
    }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("QUY_LON");
        var data = sheet.getDataRange().getValues();

        var remainingPay = parseFloat(payAmount);
        var now = new Date();

        // Ghép Tên người thu + Ngày giờ (Ví dụ: admin - 10/06/2026 07:49:10)
        var ngayGioTra = nguoiThu + " - " + now.toLocaleDateString('en-GB') + " " + now.toLocaleTimeString('en-GB');

        // Quét từ trên xuống dưới tìm các dòng đang nợ của user này
        for (var i = 1; i < data.length; i++) {
            if (remainingPay <= 0) break; // Hết tiền trả thì dừng quét

            var rowMa = data[i][1]; // Cột B: Mã thẻ
            var noHienTai = parseFloat(data[i][5]) || 0; // Cột F: Ghi nợ

            // Nếu đúng người và đang có nợ
            if (rowMa === maThe && noHienTai > 0) {
                var daNopCu = parseFloat(data[i][4]) || 0; // Cột E: Đã nộp

                if (remainingPay >= noHienTai) {
                    // TRẢ ĐỨT KHOẢN NỢ DÒNG NÀY
                    sheet.getRange(i + 1, 5).setValue(daNopCu + noHienTai); // Chuyển nợ sang Đã nộp
                    sheet.getRange(i + 1, 6).setValue(0); // Ghi nợ về 0
                    sheet.getRange(i + 1, 9).setValue(ngayGioTra); // Chốt người thu & giờ trả
                    remainingPay -= noHienTai; // Trừ đi số tiền vừa đắp vào
                } else {
                    // TRẢ MỘT PHẦN KHOẢN NỢ
                    sheet.getRange(i + 1, 5).setValue(daNopCu + remainingPay); // Cộng thêm tiền vừa trả
                    sheet.getRange(i + 1, 6).setValue(noHienTai - remainingPay); // Trừ lùi nợ
                    sheet.getRange(i + 1, 9).setValue(ngayGioTra); // Chốt người thu & giờ trả
                    remainingPay = 0; // Xài hết tiền trả
                }
            }
        }

        try {
            var nowStr = new Date().toLocaleString('vi-VN');
            var msg = '<b>🐷 CÓ NGƯỜI VỪA NỘP LỢN (TRẢ NỢ)</b>\n' +
                '👤 Người thu: <b>' + (nguoiThu || '') + '</b>\n' +
                '🧑 Người nộp (Mã Thẻ): <b>' + (maThe || '') + '</b>\n' +
                '💰 Số tiền nộp: <b>' + (parseFloat(payAmount).toLocaleString('vi-VN')) + 'K</b>\n' +
                '⏰ Thời gian: ' + nowStr;
            sendTelegramNotify(msg.replace(/<[^>]+>/g, ""));
        } catch (e) { }

        // Trả về dữ liệu mới nhất để web cập nhật ngay lập tức
        return sheet.getDataRange().getDisplayValues();
    } finally {
        lock.releaseLock();
    }
}











// ============================== 18. ĐIỂM MÙ TUẦN ==============================
function getDiemMuTuanDropdowns() {
    try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DATA");
        if (!sheet) return { success: false, message: "Không tìm thấy sheet DATA" };
        var lastRow = sheet.getLastRow();
        if (lastRow < 3) return { success: true, trams: [], toTruongs: [] };

        // Lấy cột A (Tên Trạm) và cột K (Xác Nhận / Tổ Trưởng), bắt đầu từ dòng 3
        var dataA = sheet.getRange(3, 1, lastRow - 2, 1).getValues();
        var dataK = sheet.getRange(3, 11, lastRow - 2, 1).getValues();

        var tramSet = {}, ttSet = {};
        for (var i = 0; i < dataA.length; i++) {
            var tram = String(dataA[i][0] || "").trim();
            var tt = String(dataK[i][0] || "").trim();
            if (tram) tramSet[tram] = 1;
            if (tt) ttSet[tt] = 1;
        }
        return {
            success: true,
            trams: Object.keys(tramSet).sort(),
            toTruongs: Object.keys(ttSet).sort()
        };
    } catch (e) { return { success: false, message: e.toString() }; }
}

function getDiemMuTuanData(params) {
    try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DATA");
        if (!sheet) return { success: false, message: "Không tìm thấy sheet DATA" };
        var lastRow = sheet.getLastRow();
        if (lastRow < 3) return { success: true, data: {}, lines: [], dateRange: "" };

        var fromDateStr = params.fromDate || "";
        var toDateStr = params.toDate || "";
        // filterTrams và filterTTs là mảng - rỗng = tất cả
        var filterTrams = Array.isArray(params.filterTrams) ? params.filterTrams : (params.filterTram ? [params.filterTram.trim()] : []);
        var filterTTs = Array.isArray(params.filterTTs) ? params.filterTTs : (params.filterTT ? [params.filterTT.trim()] : []);

        var fromDate = fromDateStr ? new Date(fromDateStr) : null;
        var toDate = toDateStr ? new Date(toDateStr) : null;
        if (fromDate) fromDate.setHours(0, 0, 0, 0);
        if (toDate) toDate.setHours(23, 59, 59, 999);

        // Lấy các cột cần thiết: A(0), E(4), F(5), J(9), K(10), L(11)
        // getRange(startRow, startCol, numRows, numCols) -> cột 1=A
        var numRows = lastRow - 2;
        var data = sheet.getRange(3, 1, numRows, 14).getValues();

        // Dùng Map để tổng hợp: key = tram -> lineKey -> {kt, loi}
        var pivotData = {};   // {tram: {lineKey: {kt, loi}}}
        var lineMap = {};     // {lineKey: {line, toTruong}} - thứ tự xuất hiện

        for (var i = 0; i < data.length; i++) {
            var r = data[i];
            var tram = String(r[0] || "").trim();
            if (!tram) continue;

            // Lọc theo trạm (multi-select: [] = tất cả)
            if (filterTrams.length > 0 && filterTrams.indexOf(tram) === -1) continue;

            // Lọc theo ngày - cột L = index 11
            var dateVal = r[11];
            var rowDate = (dateVal instanceof Date) ? dateVal : (dateVal ? new Date(dateVal) : null);
            if (rowDate && isNaN(rowDate.getTime())) rowDate = null;
            if (fromDate && rowDate && rowDate < fromDate) continue;
            if (toDate && rowDate && rowDate > toDate) continue;

            // Cột J = index 9 (Line), K = index 10 (Xác nhận / Tổ Trưởng)
            var line = String(r[9] || "").trim();
            var toTruong = String(r[10] || "").trim();
            if (!line && !toTruong) continue;

            // Lọc theo tổ trưởng (multi-select: [] = tất cả)
            if (filterTTs.length > 0 && filterTTs.indexOf(toTruong) === -1) continue;

            // Cột E = index 4 (Kiểm Tra), F = index 5 (Lỗi)
            var kt = parseFloat(r[4]) || 0;
            var loi = parseFloat(r[5]) || 0;

            var lineKey = line + "|||" + toTruong;

            if (!lineMap[lineKey]) {
                lineMap[lineKey] = { line: line || toTruong, toTruong: toTruong, key: lineKey };
            }
            if (!pivotData[tram]) pivotData[tram] = {};
            if (!pivotData[tram][lineKey]) pivotData[tram][lineKey] = { kt: 0, loi: 0 };
            pivotData[tram][lineKey].kt += kt;
            pivotData[tram][lineKey].loi += loi;
        }

        // Sắp xếp lines theo thứ tự xuất hiện
        var linesArr = Object.values(lineMap).sort(function (a, b) {
            return a.line.localeCompare(b.line);
        });

        // Tạo dateRange string
        var dateRange = "";
        if (fromDateStr && toDateStr) {
            var f = fromDateStr.split("-"); var t = toDateStr.split("-");
            dateRange = f[2] + "/" + f[1] + "/" + f[0] + " - " + t[2] + "/" + t[1] + "/" + t[0];
        }

        return { success: true, data: pivotData, lines: linesArr, dateRange: dateRange };
    } catch (e) { return { success: false, message: e.toString() }; }
}

function getTraCuuData() {
    try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DATA");
        if (!sheet) return { success: false, message: "Không tìm thấy sheet DATA" };

        // Dùng getValues() thay vì getDisplayValues() để tăng tốc độ gấp nhiều lần
        var data = sheet.getDataRange().getValues();
        return { success: true, data: data };
    } catch (e) { return { success: false, message: e.toString() }; }
}

function getTraCuuDropdowns() {
    try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DATA");
        if (!sheet) return { success: false, message: "Không tìm thấy sheet DATA" };

        // Lấy đúng những cột cần thiết để nhanh hơn (cột B đến N -> cột 2 đến 14)
        var lastRow = sheet.getLastRow();
        if (lastRow < 3) return { success: true, dropdowns: { mathe: [], line: [], maloi: [], model: [], ca: [] } };

        var data = sheet.getRange(3, 2, lastRow - 2, 13).getValues();
        var matheSet = {}, lineSet = {}, maloiSet = {}, modelSet = {}, caSet = {};

        for (var i = 0; i < data.length; i++) {
            var r = data[i];
            if (!r[0]) continue; // Mã thẻ ở cột 0 (vì getRange từ cột 2)
            var mt = String(r[0]).trim().toUpperCase();
            var ln = r[8] ? String(r[8]).trim().toUpperCase() : ""; // Cột J là index 8
            var ml = r[5] ? String(r[5]).trim().toUpperCase() : ""; // Cột G là index 5
            var md = r[12] ? String(r[12]).trim().toUpperCase() : ""; // Cột N là index 12
            var ca = r[11] ? String(r[11]).trim().toUpperCase() : ""; // Cột M là index 11

            if (mt) matheSet[mt] = 1;
            if (ln) lineSet[ln] = 1;
            if (ml) maloiSet[ml] = 1;
            if (md) modelSet[md] = 1;
            if (ca) caSet[ca] = 1;
        }

        return {
            success: true,
            dropdowns: {
                mathe: Object.keys(matheSet).sort(),
                line: Object.keys(lineSet).sort(),
                maloi: Object.keys(maloiSet).sort(),
                model: Object.keys(modelSet).sort(),
                ca: Object.keys(caSet).sort()
            }
        };
    } catch (e) { return { success: false, message: e.toString() }; }
}

function searchTraCuuBackend(params) {
    try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DATA");
        if (!sheet) return { success: false, message: "Không tìm thấy sheet DATA" };
        var data = sheet.getDataRange().getValues();
        var results = [];
        var isFastSearch = params.isFastSearch;
        var fastMaThes = params.fastMaThes || [];
        var vMaThes = params.vMaThes || [];
        var vLines = params.vLines || [];
        var vMaLois = params.vMaLois || [];
        var vModels = params.vModels || [];
        var vCas = params.vCas || [];
        var fromDate = params.fromDate ? new Date(params.fromDate) : null;
        var toDate = params.toDate ? new Date(params.toDate) : null;
        if (fromDate) fromDate.setHours(0, 0, 0, 0);
        if (toDate) toDate.setHours(23, 59, 59, 999);

        for (var i = 2; i < data.length; i++) {
            var r = data[i];
            if (!r || !r[1] || String(r[1]).trim() === "") continue;

            if (isFastSearch) {
                var rMaThe = String(r[1]).trim().toUpperCase();
                if (fastMaThes.indexOf(rMaThe) === -1) continue;
            } else {
                var rDateObj = (r[11] instanceof Date) ? r[11] : (r[11] ? new Date(r[11]) : null);
                if (fromDate && rDateObj && rDateObj < fromDate) continue;
                if (toDate && rDateObj && rDateObj > toDate) continue;

                var rMaThe = String(r[1]).trim().toUpperCase();
                var rLine = r[9] ? String(r[9]).trim().toUpperCase() : "";
                var rMaLoi = r[6] ? String(r[6]).trim().toUpperCase() : "";
                var rModel = r[13] ? String(r[13]).trim().toUpperCase() : "";
                var rCa = r[12] ? String(r[12]).trim().toUpperCase() : "";

                if (vMaThes.length > 0 && vMaThes.indexOf(rMaThe) === -1) continue;
                if (vLines.length > 0 && vLines.indexOf(rLine) === -1) continue;
                if (vMaLois.length > 0 && vMaLois.indexOf(rMaLoi) === -1) continue;
                if (vModels.length > 0 && vModels.indexOf(rModel) === -1) continue;
                if (vCas.length > 0 && vCas.indexOf(rCa) === -1) continue;
            }

            var dt = r[11];
            var ds = "";
            if (dt instanceof Date) {
                var dd = String(dt.getDate()).padStart(2, '0');
                var mm = String(dt.getMonth() + 1).padStart(2, '0');
                ds = dd + "/" + mm + "/" + dt.getFullYear();
            } else {
                ds = dt || "";
            }
            results.push([r[0], r[1], r[2], r[4], r[5], r[6], r[7], r[8], r[9], r[10], ds, r[12], r[13]]);
            if (results.length >= 3000) break;
        }
        return { success: true, data: results };
    } catch (e) { return { success: false, message: e.toString() }; }
}



function saveFundHoldersBackend(holders) {
    try {
        var props = PropertiesService.getScriptProperties();
        props.setProperty('FUND_HOLDERS', JSON.stringify(holders));
        incrementDataVersion();
        return { success: true };
    } catch (e) {
        return { success: false, error: e.toString() };
    }
}

// ---------------- PHÂN TÍCH DATA ----------------
// --- QUẢN LÝ QUỸ ---
function saveQuyData(dataObj) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
    } catch (e) {
        return { success: false, message: "Hệ thống đang bận do có người khác đang thao tác. Vui lòng thử lại sau vài giây!" };
    }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("QUỸ");

        var thoiGian = dataObj.thoiGian || "";
        var thangGhiQuy = dataObj.thangGhiQuy || "";
        if (!thangGhiQuy && thoiGian.indexOf("/") > -1) {
            var parts = thoiGian.split(" ")[0].split("/");
            if (parts.length >= 2) {
                thangGhiQuy = parts[1] + "/" + (parts[2].length === 2 ? "20" + parts[2] : parts[2]);
            }
        }

        var dataRow = [thoiGian, dataObj.loai, dataObj.soTien, dataObj.nguoiNopNhan, dataObj.lyDo, dataObj.nguoiThaoTac, thangGhiQuy];

        sheet.appendRow(dataRow);
        SpreadsheetApp.flush();
        try {
            var loaiIcon = dataObj.loai === 'Thu Tiền' ? '🟢' : '🔴';
            var msg = '🟢 [CTQ] Giao Dịch Quỹ Mới\n' +
                'Người tạo: ' + (dataObj.nguoiThaoTac || '') + '\n' +
                'Loại: ' + (dataObj.loai || '') + '\n' +
                'Số tiền: ' + (parseInt(dataObj.soTien || 0).toLocaleString('vi-VN')) + 'K\n' +
                'Đối tượng: ' + (dataObj.nguoiNopNhan || '') + '\n' +
                'Lý do: ' + (dataObj.lyDo || '') + '\n' +
                'Thời gian: ' + thoiGian;
            sendTelegramNotify(msg);
        } catch (e) { }
        return { success: true, newQuy: sheet.getDataRange().getDisplayValues() };
    } catch (e) {
        return { success: false, message: e.toString() };
    } finally {
        lock.releaseLock();
    }
}

function saveMultipleQuyData(dataArr) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
    } catch (e) {
        return { success: false, message: "Hệ thống đang bận do có người khác đang thao tác. Vui lòng thử lại sau vài giây!" };
    }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("QUỸ");

        var rowsToAppend = [];
        var telegramMsgs = [];

        for (var i = 0; i < dataArr.length; i++) {
            var dataObj = dataArr[i];
            var thoiGian = dataObj.thoiGian || "";
            var thangGhiQuy = dataObj.thangGhiQuy || "";
            if (!thangGhiQuy && thoiGian.indexOf("/") > -1) {
                var parts = thoiGian.split(" ")[0].split("/");
                if (parts.length >= 2) {
                    thangGhiQuy = parts[1] + "/" + (parts[2].length === 2 ? "20" + parts[2] : parts[2]);
                }
            }
            rowsToAppend.push([thoiGian, dataObj.loai, dataObj.soTien, dataObj.nguoiNopNhan, dataObj.lyDo, dataObj.nguoiThaoTac, thangGhiQuy]);
            var loaiIcon = dataObj.loai === 'Thu Tiền' ? '🟢' : '🔴';
            var msg = '🟢 [CTQ] Giao Dịch Quỹ Mới\n' +
                'Người tạo: ' + (dataObj.nguoiThaoTac || '') + '\n' +
                'Loại: ' + (dataObj.loai || '') + '\n' +
                'Số tiền: ' + (parseInt(dataObj.soTien || 0).toLocaleString('vi-VN')) + 'K\n' +
                'Đối tượng: ' + (dataObj.nguoiNopNhan || '') + '\n' +
                'Lý do: ' + (dataObj.lyDo || '') + '\n' +
                'Thời gian: ' + thoiGian;
            telegramMsgs.push(msg);
        }

        if (rowsToAppend.length > 0) {
            var lastRow = sheet.getLastRow();
            sheet.getRange(lastRow + 1, 1, rowsToAppend.length, 7).setValues(rowsToAppend);
            SpreadsheetApp.flush();

            try {
                for (var j = 0; j < telegramMsgs.length; j++) {
                    sendTelegramNotify(telegramMsgs[j]);
                }
            } catch (e) { }
        }

        return { success: true, newQuy: sheet.getDataRange().getDisplayValues() };
    } catch (e) {
        return { success: false, message: e.toString() };
    } finally {
        lock.releaseLock();
    }
}

function xoaQuyData(rows, currentUser) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
    } catch (e) {
        return { success: false, message: "Hệ thống đang bận do có người khác đang thao tác. Vui lòng thử lại sau vài giây!" };
    }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("QUỸ");
        if (!sheet) return { success: false, message: "Không tìm thấy sheet QUỸ" };

        var data = sheet.getDataRange().getDisplayValues();
        var deletedDetails = [];

        if (rows && rows.length > 0) {
            rows.sort(function (a, b) { return b - a });
            for (var i = 0; i < rows.length; i++) {
                var r = rows[i];
                if (data[r - 1]) {
                    var d = data[r - 1];
                    deletedDetails.push("Loại: " + d[1] + " | Tiền: " + d[2] + " | Người: " + d[3] + " | Lý do: " + d[4]);
                }
                sheet.deleteRow(r);
            }
        }
        logAction("Xoá Dữ Liệu Quỹ", currentUser, "Đã xoá " + rows.length + " dòng dữ liệu quỹ:\\n" + deletedDetails.join("\\n"));
        SpreadsheetApp.flush();
        return { success: true, newQuy: sheet.getDataRange().getDisplayValues() };
    } catch (e) {
        return { success: false, message: e.toString() };
    } finally {
        lock.releaseLock();
    }
}






// --- WEBHOOK LOGGER ---
function logAction(actionName, currentUser, details) {
    try {
        var msg = "[CTQ] " + actionName + "\n" +
            "Người thao tác: " + (currentUser || "Khách") + "\n";
        if (details) {
            msg += "Chi tiết:\n" + details + "\n";
        }
        msg += "Thời gian: " + new Date().toLocaleString('vi-VN');

        // Bỏ qua các hành động đã có thông báo riêng biệt để tránh trùng lặp Telegram
        var skipTelegram = ["Lưu Nhập Thẻ", "Lưu Nhập Thẻ -> Hệ Thống", "Đăng Nhập", "Đăng Xuất", "Yêu Cầu Cấp Thẻ"];
        if (skipTelegram.indexOf(actionName) === -1) {
            sendTelegramNotify(msg);
        }
    } catch (e) { }
}

function getSpreadsheet() {
    var ssId = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
    if (ssId) {
        try { return SpreadsheetApp.openById(ssId); } catch (e) { }
    }
    return SpreadsheetApp.getActiveSpreadsheet();
}

// --- XỬ LÝ TIN NHẮN TỪ WEBHOOK ---
function doPost(e) {
    try {
        var rawData = e ? (e.postData ? e.postData.contents : "No postData") : "No e";
        try {
            var ss = getSpreadsheet();
            if (ss) {
                var logSheet = ss.getSheetByName("WEBHOOK_LOG");
                if (!logSheet) {
                    logSheet = ss.insertSheet("WEBHOOK_LOG");
                    logSheet.appendRow(["Time", "Payload"]);
                }
                logSheet.appendRow([new Date(), rawData]);
            }
        } catch (logErr) { }

        if (!e || !e.postData || !e.postData.contents) return ContentService.createTextOutput("OK");
        var data = JSON.parse(e.postData.contents);
        if (!data.message || !data.message.text) return ContentService.createTextOutput("OK");

        var chatId = null;
        if (data.message.chat && data.message.chat.id) chatId = data.message.chat.id;
        else if (data.message.from && data.message.from.id) chatId = data.message.from.id;
        else if (data.chat && data.chat.id) chatId = data.chat.id;

        if (!chatId) return ContentService.createTextOutput("OK");

        var originalText = data.message.text || "";
        var text = originalText.toLowerCase().trim();
        // Loại bỏ dấu gạch chéo nếu người dùng có gõ
        text = text.replace(/\//g, "");

        var isReplyChat = false;
        var replyReceiver = "";
        if (data.message.reply_to_message && data.message.reply_to_message.text) {
            var repText = data.message.reply_to_message.text;
            if (repText.indexOf("💬 TIN NHẮN CHAT TỪ ") > -1) {
                var lines = repText.split("\n");
                replyReceiver = lines[0].replace("💬 TIN NHẮN CHAT TỪ ", "").trim();
                isReplyChat = true;
            }
        }

        var replyMessage = "🤖 Chào bạn, tôi là AI Trợ Lý của hệ thống CTQ. Bạn cần tôi giúp gì?\\nBạn có thể thử các lệnh:\\n👉 <b>nhân viên</b>\\n👉 <b>cấp thẻ</b>\\n👉 <b>quỹ lợn</b>\\n👉 <b>tổng quỹ</b>";

        if (text.startsWith("debug")) {
            replyMessage = "🛠 DEBUG TEXT: [" + originalText + "] | [" + text + "]";
        } else if (text === "id") {
            replyMessage = "🔑 Mã ID Telegram của bạn là:\n" + chatId;
        } else if (text === "start" || text === "help" || text === "hi" || text === "hello" || text === "chào" || text === "chao") {
            replyMessage = "🤖 <b>PHIÊN BẢN BOT MỚI V2 XIN CHÀO!</b>\n\nBạn có thể dùng các lệnh nhanh:\n" +
                "👉 <b>1</b>: Điểm mù chưa xác nhận\n" +
                "👉 <b>2 [ngày]</b>: Tổng hợp lỗi điểm mù theo ngày\n" +
                "👉 <b>3</b> (hoặc <b>cấp thẻ</b>): Yêu cầu thẻ chờ xử lý\n" +
                "👉 <b>4</b> (hoặc <b>quỹ lợn</b>): Nợ quỹ lợn\n" +
                "👉 <b>5</b> (hoặc <b>tổng quỹ</b>): Tồn quỹ hiện tại\n\n" +
                "Và các lệnh khác:\n" +
                "- <b>tìm [mã thẻ/tên]</b>: Tra cứu NV\n" +
                "- <b>ghi nợ [mã thẻ] [tiền] [lý do]</b>: Phạt lợn\n" +
                "- <b>nộp lợn [mã thẻ] [tiền]</b>: Thu nợ\n" +
                "- <b>duyệt thẻ [mã thẻ]</b>: Duyệt thẻ chờ\n" +
                "- <b>quỹ tháng [X]</b>: Báo cáo tháng\n" +
                "- <b>nhân viên</b>: Đếm số NV\n" +
                "- <b>load</b>: Tóm tắt hệ thống";
        } else if (text === "1") {
            var ss = getSpreadsheet();
            var sheet = ss ? ss.getSheetByName("DATA") : null;
            var sheetData = sheet ? sheet.getDataRange().getDisplayValues() : [];
            var cxnList = [];
            for (var i = 2; i < sheetData.length; i++) {
                var status = String(sheetData[i][14] || "");
                if (status.indexOf("Chưa Xác Nhận") > -1) {
                    var line = sheetData[i][9] || "N/A";
                    var the = sheetData[i][1] || "N/A";
                    var maLoi = sheetData[i][6] || "N/A";
                    var tenLoi = sheetData[i][7] || "N/A";
                    var toTruong = sheetData[i][10] || "N/A"; // Cột K
                    var nguoiNhap = sheetData[i][8] || "N/A"; // Cột I
                    var thoiGianNhap = sheetData[i][0] || "";

                    var minutes = 0;
                    if (thoiGianNhap) {
                        var parts = thoiGianNhap.split(" ");
                        if (parts.length === 2) {
                            var dParts = parts[0].split("/");
                            var tParts = parts[1].split(":");
                            if (dParts.length >= 3 && tParts.length >= 2) {
                                // Xử lý chuẩn múi giờ GMT+7 để không bị lệch thời gian máy chủ
                                var isoStr = dParts[2] + "-" + dParts[1] + "-" + dParts[0] + "T" + tParts[0] + ":" + tParts[1] + ":" + (tParts[2] || "00") + "+07:00";
                                var d = new Date(isoStr);
                                var diffMs = Date.now() - d.getTime();
                                minutes = Math.floor(diffMs / 60000);
                            }
                        }
                    }
                    var tgStr = "Vừa xong.";
                    if (minutes > 0) {
                        if (minutes >= 60) {
                            var hrs = Math.floor(minutes / 60);
                            var mins = minutes % 60;
                            tgStr = hrs + " Giờ " + mins + " Phút";
                        } else {
                            tgStr = minutes + " Phút";
                        }
                    }

                    cxnList.push(
                        "Line : " + line + "\n" +
                        "Mã Thẻ :" + the + "\n" +
                        "Mã Lỗi: " + maLoi + "\n" +
                        "Tên Lỗi: " + tenLoi + "\n" +
                        "Tổ Trưởng: " + toTruong + "\n" +
                        "Thời Gian Chưa Xác Nhận: " + tgStr + "\n" +
                        "Trạng Thái: Chưa Xác Nhận\n" +
                        "Người Nhập: " + nguoiNhap + "\n------------------------"
                    );
                }
            }
            if (cxnList.length > 0) {
                replyMessage = cxnList.join("\n");
                if (replyMessage.length > 3500) replyMessage = replyMessage.substring(0, 3500) + "\n... (Còn nữa)";
            } else {
                replyMessage = "Hiện Không Có Lỗi Nào Chưa Xác Nhận.";
            }
        } else if (text.startsWith("2")) {
            var dateInput = text.replace("2", "").trim();
            var targetDateStr = dateInput;
            if (!targetDateStr) {
                var now = new Date();
                targetDateStr = ('0' + now.getDate()).slice(-2) + '/' + ('0' + (now.getMonth() + 1)).slice(-2) + '/' + now.getFullYear();
            } else {
                var parts = targetDateStr.split("/");
                if (parts.length === 2) {
                    targetDateStr = ('0' + parts[0]).slice(-2) + '/' + ('0' + parts[1]).slice(-2) + '/' + new Date().getFullYear();
                } else if (parts.length === 3) {
                    var y = parts[2].length === 2 ? "20" + parts[2] : parts[2];
                    targetDateStr = ('0' + parts[0]).slice(-2) + '/' + ('0' + parts[1]).slice(-2) + '/' + y;
                }
            }
            var ss = getSpreadsheet();
            var sheet = ss ? ss.getSheetByName("DATA") : null;
            var sheetData = sheet ? sheet.getDataRange().getDisplayValues() : [];
            var totalLoi = 0;
            var caNgay = 0;
            var caDem = 0;
            for (var i = 2; i < sheetData.length; i++) {
                var rowDate = String(sheetData[i][11] || "").trim();
                if (rowDate === targetDateStr) {
                    var maLoi = String(sheetData[i][6] || "").trim();
                    if (maLoi === "CTQ-00" || maLoi === "") continue;
                    var soLoi = parseInt(sheetData[i][5] || "0");
                    totalLoi += soLoi;
                    var ca = String(sheetData[i][12] || "").trim().toLowerCase();
                    if (ca === "ngày" || ca === "ngay") caNgay += soLoi;
                    else if (ca === "đêm" || ca === "dem") caDem += soLoi;
                }
            }
            if (totalLoi > 0) {
                replyMessage = "Ngày " + targetDateStr + " Có " + totalLoi + " lỗi\n" +
                    "Ca Ngày: " + caNgay + " Lỗi\n" +
                    "Ca Đêm: " + caDem + " Lỗi";
            } else {
                replyMessage = "Ngày " + targetDateStr + " Không Có Lỗi Nào.";
            }
        } else if (text.startsWith("tìm ")) {
            var keyword = text.replace("tìm ", "").trim().toLowerCase();
            var ss = getSpreadsheet();
            var sheet = ss ? ss.getSheetByName("THÔNG TIN THẺ") : null;
            var d = sheet ? sheet.getDataRange().getDisplayValues() : [];
            var found = [];
            for (var i = 1; i < d.length; i++) {
                if (String(d[i][3]).toLowerCase().indexOf(keyword) > -1 || String(d[i][4]).toLowerCase().indexOf(keyword) > -1) {
                    found.push("- <b>" + d[i][3] + "</b> | " + d[i][4] + " | " + d[i][2]);
                }
            }
            if (found.length > 0) {
                replyMessage = "🔍 <b>KẾT QUẢ TÌM KIẾM:</b>\n" + found.join("\n");
            } else {
                replyMessage = "❌ Không tìm thấy nhân viên nào khớp với '" + keyword + "'";
            }
        } else if (text.startsWith("ghi nợ ") || text.startsWith("ghi no ")) {
            var parts = text.replace("ghi nợ ", "").replace("ghi no ", "").trim().split(" ");
            if (parts.length >= 3) {
                var maThe = parts[0];
                var tien = parseInt(parts[1]) || 0;
                var lyDo = parts.slice(2).join(" ");
                var now = new Date().toLocaleString('vi-VN');
                var dRow = [now, maThe, lyDo, 0, tien];
                var ss = getSpreadsheet();
                if (ss) ss.getSheetByName("QUY_LON").appendRow(dRow);
                replyMessage = "✅ <b>GHI NỢ THÀNH CÔNG</b>\nMã thẻ: " + maThe + "\nSố tiền: " + tien + "K\nLý do: " + lyDo;
                logAction("Ghi Nợ (Bot)", "Telegram Bot", "Mã thẻ: " + maThe + " | Nợ: " + tien + "K");
            } else {
                replyMessage = "⚠️ Sai cú pháp! Vui lòng dùng: ghi nợ [mã thẻ] [số tiền K] [lý do]";
            }
        } else if (text.startsWith("nộp lợn ") || text.startsWith("nop lon ")) {
            var parts = text.replace("nộp lợn ", "").replace("nop lon ", "").trim().split(" ");
            if (parts.length >= 2) {
                var maThe = parts[0];
                var tien = parseInt(parts[1]) || 0;
                try {
                    processTraNoThongMinh(maThe, tien, "Telegram Bot");
                    replyMessage = "✅ <b>NỘP LỢN THÀNH CÔNG</b>\nMã thẻ: " + maThe + "\nSố tiền nộp: " + tien + "K";
                } catch (e) {
                    replyMessage = "❌ Lỗi: " + e.message;
                }
            } else {
                replyMessage = "⚠️ Sai cú pháp! Vui lòng dùng: nộp lợn [mã thẻ] [số tiền K]";
            }
        } else if (text.startsWith("duyệt thẻ ") || text.startsWith("duyet the ")) {
            var maThe = text.replace("duyệt thẻ ", "").replace("duyet the ", "").trim();
            if (maThe) {
                var res = xoaCapTheDungChung([maThe], "Telegram Bot");
                if (res && res.success) {
                    replyMessage = "✅ <b>ĐÃ DUYỆT THẺ</b>\nMã thẻ: " + maThe + " đã được xác nhận cấp/in thẻ và xoá khỏi chờ.";
                } else {
                    replyMessage = "❌ Lỗi khi duyệt thẻ.";
                }
            } else {
                replyMessage = "⚠️ Sai cú pháp! Vui lòng dùng: duyệt thẻ [mã thẻ]";
            }
        } else if (text.indexOf("nhân viên") > -1 || text.indexOf("nhan vien") > -1) {
            var ss = getSpreadsheet();
            var sheet = ss ? ss.getSheetByName("THÔNG TIN THẺ") : null;
            var count = sheet ? Math.max(0, sheet.getLastRow() - 1) : 0;
            replyMessage = "👥 <b>TỔNG SỐ NHÂN VIÊN:</b> " + count + " người đang hoạt động.";
        } else if (text === "3" || text.indexOf("cấp thẻ") > -1 || text.indexOf("cap the") > -1) {
            var ss = getSpreadsheet();
            var sheet = ss ? ss.getSheetByName("CẤP THẺ") : null;
            var data = sheet ? sheet.getDataRange().getDisplayValues() : [];
            var list = [];
            for (var i = 1; i < data.length; i++) {
                var bp = data[i][0] || "";
                var model = data[i][1] || "";
                var line = data[i][2] || "";
                var the = data[i][3] || "";
                var ten = data[i][4] || "";
                var lyDo = data[i][5] || "";
                var tt = data[i][6] || "";
                var tram = data[i][7] || "";
                var str = i + ": " + bp + "," + line + "," + model + "," + the + "," + ten + "," + lyDo + "," + tt + "," + tram;
                list.push(str);
            }
            if (list.length > 0) {
                replyMessage = list.join("\n");
                if (replyMessage.length > 3500) replyMessage = replyMessage.substring(0, 3500) + "\n...";
            } else {
                replyMessage = "Hiện Không Có Yêu Cầu Cấp Thẻ Nào.";
            }
        } else if (text === "4" || text.indexOf("quỹ lợn") > -1 || text.indexOf("quy lon") > -1) {
            var ss = getSpreadsheet();
            var sheet = ss ? ss.getSheetByName("QUY_LON") : null;
            var qData = sheet ? sheet.getDataRange().getDisplayValues() : [];

            var tSheet = ss ? ss.getSheetByName("THÔNG TIN THẺ") : null;
            var tData = tSheet ? tSheet.getDataRange().getDisplayValues() : [];
            var nameMap = {};
            for (var j = 1; j < tData.length; j++) {
                nameMap[String(tData[j][3]).toUpperCase()] = tData[j][4];
            }

            var list = [];
            var totalThu = 0;
            var stt = 1;
            for (var i = 1; i < qData.length; i++) {
                var daThu = parseInt(String(qData[i][3]).replace(/\D/g, '')) || 0;
                var no = parseInt(String(qData[i][4]).replace(/\D/g, '')) || 0;
                totalThu += daThu;
                if (no > 0) {
                    var ngay = String(qData[i][0]).split(" ")[0] || "";
                    var the = String(qData[i][1]).toUpperCase() || "";
                    var lyDo = qData[i][2] || "";
                    var ten = nameMap[the] || "Không Rõ Tên";
                    list.push(stt + ": NỢ " + no + "k Gồm 1 " + ten + " " + lyDo + " ngày " + ngay + " " + no + "k");
                    stt++;
                }
            }
            replyMessage = list.join("\n") + "\nTổng Đã Đóng Lợn Có : " + totalThu.toLocaleString('vi-VN') + "K";
        } else if (text.startsWith("5") || text.indexOf("tổng quỹ") > -1 || text.indexOf("quỹ tháng") > -1) {
            var ss = getSpreadsheet();
            var sheet = ss ? ss.getSheetByName("QUỸ") : null;
            var qData = sheet ? sheet.getDataRange().getDisplayValues() : [];

            var targetMonth = "";
            var match = text.match(/(?:t|tháng|thang)[\s]*(\d+)/i);
            if (match) targetMonth = match[1];

            var thu = 0, chi = 0;
            var paidUsers = [];
            for (var i = 1; i < qData.length; i++) {
                var dateStr = String(qData[i][0] || "");
                var matchMonth = false;
                if (targetMonth === "") matchMonth = true;
                else if (dateStr.indexOf("/" + targetMonth + "/") > -1 || dateStr.startsWith(targetMonth + "/")) matchMonth = true;
                else if (targetMonth.length === 1 && (dateStr.indexOf("/0" + targetMonth + "/") > -1 || dateStr.startsWith("0" + targetMonth + "/"))) matchMonth = true;

                if (matchMonth) {
                    var rThu = parseFloat(String(qData[i][3]).replace(/,/g, '')) || 0;
                    var rChi = parseFloat(String(qData[i][4]).replace(/,/g, '')) || 0;
                    thu += rThu;
                    chi += rChi;
                    if (rThu > 0) {
                        paidUsers.push(String(qData[i][1]).toUpperCase());
                        paidUsers.push(String(qData[i][5]).toUpperCase());
                    }
                }
            }
            var ton = thu - chi;

            if (targetMonth !== "") {
                var tSheet = ss ? ss.getSheetByName("THÔNG TIN THẺ") : null;
                var tData = tSheet ? tSheet.getDataRange().getDisplayValues() : [];
                var unpaid = [];
                for (var j = 1; j < tData.length; j++) {
                    var maThe = String(tData[j][3]).toUpperCase();
                    var ten = String(tData[j][4]);
                    if (maThe && maThe !== "MÃ THẺ") {
                        var hasPaid = false;
                        for (var k = 0; k < paidUsers.length; k++) {
                            if (paidUsers[k].indexOf(maThe) > -1 || (ten && paidUsers[k].indexOf(ten.toUpperCase()) > -1)) {
                                hasPaid = true; break;
                            }
                        }
                        if (!hasPaid) unpaid.push(ten + " (" + maThe + ")");
                    }
                }

                replyMessage = "Tổng Quỹ Tháng " + targetMonth + " Có " + thu.toLocaleString('vi-VN') + "K Tiền\n" +
                    "Còn " + unpaid.length + " Người Chưa Đóng Quỹ\n" +
                    "Tổng Chi Tháng " + targetMonth + ": " + chi.toLocaleString('vi-VN') + "K\n" +
                    "Tổng Thu Tháng " + targetMonth + ": " + thu.toLocaleString('vi-VN') + "K";
            } else {
                replyMessage = "Tổng Qũy Hiện Tại : " + thu.toLocaleString('vi-VN') + "K\n" +
                    "TỔNG CHI TIỀN: " + chi.toLocaleString('vi-VN') + "K\n" +
                    "TOTAL Hiện Tại : " + ton.toLocaleString('vi-VN') + "k";
            }
        } else if (text.startsWith("7 ")) {
            var maThe = text.replace("7", "").trim().toUpperCase();
            var ss = getSpreadsheet();
            var sheet = ss ? ss.getSheetByName("THÔNG TIN THẺ") : null;
            var d = sheet ? sheet.getDataRange().getDisplayValues() : [];
            var foundRow = null;
            var headers = d.length > 0 ? d[0] : [];
            for (var i = 1; i < d.length; i++) {
                if (String(d[i][3]).toUpperCase() === maThe) {
                    foundRow = d[i];
                    break;
                }
            }
            if (foundRow) {
                var info = [];
                for (var c = 0; c < 18 && c < headers.length; c++) {
                    info.push(headers[c] + ": " + foundRow[c]);
                }
                replyMessage = "Thông Tin Thẻ " + maThe + ":\n" + info.join("\n");
            } else {
                replyMessage = "Không Tìm Thấy Thông Tin Cho Mã Thẻ: " + maThe;
            }
        } else if (text === "load" || text === "tóm tắt" || text === "tom tat" || text === "hệ thống" || text === "he thong") {
            var ss = getSpreadsheet();
            if (!ss) {
                replyMessage = "❌ Lỗi: Không thể kết nối với Spreadsheet.";
            } else {
                var sheetNV = ss.getSheetByName("THÔNG TIN THẺ");
                var countNV = sheetNV ? Math.max(0, sheetNV.getLastRow() - 1) : 0;

                var sheetCT = ss.getSheetByName("CẤP THẺ");
                var countCT = sheetCT ? Math.max(0, sheetCT.getLastRow() - 1) : 0;

                var sheetQL = ss.getSheetByName("QUY_LON");
                var qlData = sheetQL ? sheetQL.getDataRange().getDisplayValues() : [];
                var tongNo = 0;
                for (var i = 1; i < qlData.length; i++) {
                    tongNo += parseInt(String(qlData[i][4]).replace(/\D/g, '')) || 0;
                }

                var sheetQuy = ss.getSheetByName("QUỸ");
                var qData = sheetQuy ? sheetQuy.getDataRange().getDisplayValues() : [];
                var thu = 0, chi = 0;
                for (var i = 1; i < qData.length; i++) {
                    if (qData[i][3]) thu += parseFloat(String(qData[i][3]).replace(/,/g, '')) || 0;
                    if (qData[i][4]) chi += parseFloat(String(qData[i][4]).replace(/,/g, '')) || 0;
                }
                var ton = thu - chi;

                replyMessage = "📊 <b>TÓM TẮT HỆ THỐNG:</b>\n\n" +
                    "👥 Nhân viên: " + countNV + " người\n" +
                    "💳 Cấp thẻ chờ: " + countCT + " yêu cầu\n" +
                    "💰 Tồn quỹ: " + ton.toLocaleString('vi-VN') + "K\n" +
                    "🐷 Nợ quỹ lợn: " + tongNo.toLocaleString('vi-VN') + "K";
            }
        } else {
            // Không phản hồi nếu không phải là lệnh hợp lệ
        }

        if (!replyMessage) replyMessage = "⚠️ Không hiểu lệnh. Gõ 'hi' để xem trợ giúp.";

        // Xoá HTML tag khi gửi Telegram vì Telegram không hỗ trợ HTML
        var isTelegramId = !!chatId;
        if (isTelegramId) {
            var plainMsg = replyMessage.replace(/<[^>]+>/g, "");
            sendTelegramNotify(plainMsg, chatId);
        }
        return ContentService.createTextOutput("OK");
    } catch (errObj) {
        try {
            var isTelegramId = !!chatId;
            if (isTelegramId && chatId) {
                sendTelegramNotify("❌ Lỗi Code Webhook: " + errObj.toString(), chatId);
            }
        } catch (innerErr) { }
        // LUÔN LUÔN TRẢ VỀ OK ĐỂ BOT NGỪNG RETRY (Ngừng spam tin nhắn cũ)
        return ContentService.createTextOutput("OK");
    }
}



// --- KIỂM TRA LỖI NHÂN VIÊN ---
function getKiemTraData(fromDate, toDate, soLoiTarget) {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var dataSheet = ss.getSheetByName("DATA");
        if (!dataSheet) return [];
        var lastRow = dataSheet.getLastRow();
        if (lastRow < 2) return [];

        // Lấy 13 cột từ A đến M
        var data = dataSheet.getRange(2, 1, lastRow - 1, 13).getDisplayValues();

        var fDate = new Date(fromDate); fDate.setHours(0, 0, 0, 0);
        var tDate = new Date(toDate); tDate.setHours(23, 59, 59, 999);

        var validRows = [];
        var countByMaThe = {};

        // 1. Lọc theo khoảng thời gian và đếm số lần xuất hiện có lỗi của MÃ THẺ
        for (var i = 0; i < data.length; i++) {
            var row = data[i];
            var loi = parseFloat(row[5]) || 0; // Cột F (LỖI)

            if (loi < 1) continue; // Điều kiện là phải cột F hiện số 1 trở lên

            var rowDate = row[11]; // Cột L (NGÀY)
            var rDate;
            var parts = String(rowDate).split('/');
            if (parts.length === 3) {
                rDate = new Date(parts[2], parts[1] - 1, parts[0]);
            } else {
                rDate = new Date(rowDate);
            }

            if (!isNaN(rDate.getTime()) && rDate >= fDate && rDate <= tDate) {
                validRows.push(row);
                var maThe = String(row[1]).trim().toUpperCase(); // Cột B (MÃ THẺ)
                if (maThe) {
                    if (!countByMaThe[maThe]) countByMaThe[maThe] = 0;
                    countByMaThe[maThe]++;
                }
            }
        }

        // 2. Lọc lại các row có MÃ THẺ xuất hiện đúng 'soLoiTarget' lần (hoặc tất cả nếu là 'ALL')
        var result = [];
        for (var j = 0; j < validRows.length; j++) {
            var mThe = String(validRows[j][1]).trim().toUpperCase();
            if (soLoiTarget === 'ALL' || countByMaThe[mThe] === soLoiTarget) {
                result.push(validRows[j]);
            }
        }

        return result;
    } catch (e) {
        return [];
    }
}

// --- BÁO CÁO GIÁO VIÊN ---
function getGiaoVienData(fromDate, toDate) {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var dataSheet = ss.getSheetByName("DATA");
        if (!dataSheet) return { success: false, data: [] };
        var lastRow = dataSheet.getLastRow();
        if (lastRow < 2) return { success: true, data: [] };

        var data = dataSheet.getRange(2, 1, lastRow - 1, 12).getDisplayValues();

        var fDate = new Date(fromDate); fDate.setHours(0, 0, 0, 0);
        var tDate = new Date(toDate); tDate.setHours(23, 59, 59, 999);

        var groupedByGV = {};

        for (var i = 0; i < data.length; i++) {
            var row = data[i];
            var rowDate = row[11]; // Cột L (NGAY)
            if (!rowDate) continue;

            var rDate;
            var parts = String(rowDate).split('/');
            if (parts.length === 3) {
                rDate = new Date(parts[2], parts[1] - 1, parts[0]);
            } else {
                rDate = new Date(rowDate);
            }

            if (!isNaN(rDate.getTime()) && rDate >= fDate && rDate <= tDate) {
                var gv = String(row[8]).trim(); // Cột I (Giáo Viên)
                if (!gv || gv === "" || gv === "-Chọn-") gv = "Chưa Cập Nhật";

                var testStr = String(row[4]).replace(/,/g, ''); // Cột E (Lần Test)
                var lanTest = parseFloat(testStr) || 0;

                var loiStr = String(row[5]).replace(/,/g, ''); // Cột F (Số Lỗi)
                var numLoi = parseFloat(loiStr) || 0;
                var coLoi = (numLoi >= 1) ? 1 : 0;

                if (!groupedByGV[gv]) {
                    groupedByGV[gv] = {
                        name: gv,
                        tongTest: 0,
                        tongLoi: 0
                    };
                }

                groupedByGV[gv].tongTest += lanTest;
                groupedByGV[gv].tongLoi += coLoi;
            }
        }

        var result = [];
        for (var k in groupedByGV) {
            var item = groupedByGV[k];
            var rate = (item.tongTest > 0) ? (item.tongLoi / item.tongTest * 100) : 0;
            result.push({
                name: item.name,
                tongTest: item.tongTest,
                tongLoi: item.tongLoi,
                rate: parseFloat(rate.toFixed(2))
            });
        }

        result.sort(function (a, b) {
            if (b.tongLoi !== a.tongLoi) return b.tongLoi - a.tongLoi;
            return b.tongTest - a.tongTest;
        });

        return { success: true, data: result };
    } catch (e) {
        return { success: false, message: e.message };
    }
}

// --- XẾP HẠNG TỔ TRƯỞNG ---
function getXepHangData(fromDate, toDate) {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var dataSheet = ss.getSheetByName("DATA");
        if (!dataSheet) return { success: false, data: [] };
        var lastRow = dataSheet.getLastRow();
        if (lastRow < 2) return { success: true, data: [] };

        // Lấy dữ liệu từ cột A đến M (ít nhất 12 cột để có L)
        var data = dataSheet.getRange(2, 1, lastRow - 1, 12).getDisplayValues();

        var fDate = new Date(fromDate); fDate.setHours(0, 0, 0, 0);
        var tDate = new Date(toDate); tDate.setHours(23, 59, 59, 999);

        var groupedByToTruong = {};

        // 1. Lọc và nhóm theo Tổ Trưởng
        for (var i = 0; i < data.length; i++) {
            var row = data[i];
            var loi = parseFloat(row[5]) || 0; // Cột F (LỖI)

            if (loi < 1) continue; // Cột F phải có số 1 trở lên

            var rowDate = row[11]; // Cột L (NGÀY)
            var rDate;
            var parts = String(rowDate).split('/');
            if (parts.length === 3) {
                rDate = new Date(parts[2], parts[1] - 1, parts[0]);
            } else {
                rDate = new Date(rowDate);
            }

            if (!isNaN(rDate.getTime()) && rDate >= fDate && rDate <= tDate) {
                var toTruong = String(row[10]).trim() || "Chưa Cập Nhật"; // Cột K
                var maThe = String(row[1]).trim(); // Cột B
                var hoTen = String(row[2]).trim(); // Cột C
                var tenLoi = String(row[7]).trim(); // Cột H

                if (!groupedByToTruong[toTruong]) {
                    groupedByToTruong[toTruong] = {
                        name: toTruong,
                        totalErrors: 0,
                        rows: []
                    };
                }
                groupedByToTruong[toTruong].totalErrors += loi;
                groupedByToTruong[toTruong].rows.push({
                    toTruong: toTruong,
                    maThe: maThe,
                    hoTen: hoTen,
                    soLoi: loi,
                    tenLoi: tenLoi,
                    ngay: rowDate
                });
            }
        }

        // 2. Chuyển object thành mảng và sắp xếp theo totalErrors giảm dần
        var sortedGroups = Object.values(groupedByToTruong);
        sortedGroups.sort(function (a, b) {
            return b.totalErrors - a.totalErrors;
        });

        // 3. Trải phẳng mảng để trả về
        var finalData = [];
        for (var j = 0; j < sortedGroups.length; j++) {
            var gRows = sortedGroups[j].rows;
            for (var k = 0; k < gRows.length; k++) {
                finalData.push(gRows[k]);
            }
        }

        return { success: true, data: finalData };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function getGiaoVienData(fromDate, toDate) {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var dataSheet = ss.getSheetByName("DATA");
        if (!dataSheet) return { success: false, message: "Không tìm thấy sheet DATA" };

        var lastRow = dataSheet.getLastRow();
        if (lastRow < 2) return { success: true, data: [] };

        var fDate, tDate;
        if (fromDate && toDate) {
            var dp1 = fromDate.split('-'); fDate = new Date(dp1[0], dp1[1] - 1, dp1[2]); fDate.setHours(0, 0, 0, 0);
            var dp2 = toDate.split('-'); tDate = new Date(dp2[0], dp2[1] - 1, dp2[2]); tDate.setHours(23, 59, 59, 999);
        }

        // Lấy dữ liệu bỏ qua header
        var data = dataSheet.getRange(2, 1, lastRow - 1, 14).getDisplayValues();
        var grouped = {};

        for (var i = 0; i < data.length; i++) {
            var row = data[i];

            var rowDate = row[11]; // Cột L (NGÀY) - index 11
            var rDate;
            var parts = String(rowDate).split('/');
            if (parts.length === 3) {
                rDate = new Date(parts[2], parts[1] - 1, parts[0]);
            } else {
                rDate = new Date(rowDate);
            }

            if (fDate && tDate) {
                if (isNaN(rDate.getTime()) || rDate < fDate || rDate > tDate) continue;
            }

            var giaoVien = String(row[8]).trim() || "N/A"; // Cột I (THỰC HIỆN) - index 8
            var lanTest = parseFloat(row[4]) || 0;         // Cột E (LẦN TEST) - index 4
            var loi = parseFloat(row[5]) || 0;             // Cột F (LỖI) - index 5

            // Những số lỗi từ 1 trở lên
            if (loi < 1) loi = 0;

            if (!grouped[giaoVien]) {
                grouped[giaoVien] = {
                    name: giaoVien,
                    totalTest: 0,
                    totalLoi: 0
                };
            }
            grouped[giaoVien].totalTest += lanTest;
            grouped[giaoVien].totalLoi += loi;
        }

        var resultData = Object.values(grouped);

        // Sắp xếp: test nhiều nhất và bắt được nhiều lỗi nhất -> Sort by totalTest DESC, then totalLoi DESC
        resultData.sort(function (a, b) {
            if (b.totalTest !== a.totalTest) {
                return b.totalTest - a.totalTest;
            }
            return b.totalLoi - a.totalLoi;
        });

        // Tính tỷ lệ %
        for (var i = 0; i < resultData.length; i++) {
            var item = resultData[i];
            item.rate = item.totalTest > 0 ? ((item.totalLoi / item.totalTest) * 100).toFixed(2) + '%' : '0.00%';
        }

        return { success: true, data: resultData };

    } catch (e) {
        return { success: false, message: e.toString() };
    }
}


function loadQuaySoData() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("DB_QUAYSO");
    if (!sheet) return null;
    var data = sheet.getDataRange().getValues();
    var jsonStr = "";
    for (var i = 0; i < data.length; i++) {
        if (data[i][0]) jsonStr += data[i][0];
    }
    return jsonStr;
}

function saveQuaySoData(jsonStr) {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("DB_QUAYSO");
        if (!sheet) {
            sheet = ss.insertSheet("DB_QUAYSO");
            sheet.hideSheet();
        }
        var chunks = [];
        for (var i = 0; i < jsonStr.length; i += 45000) {
            chunks.push([jsonStr.substring(i, i + 45000)]);
        }
        sheet.clear();
        if (chunks.length > 0) {
            sheet.getRange(1, 1, chunks.length, 1).setValues(chunks);
        }
        return true;
    } catch (e) {
        return false;
    }
}

// ===========================================
// QUAY SO ONLINE DATA
// ===========================================
function saveQuaySoData(jsonData) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Data_QuaySo');
    if (!sheet) {
        sheet = ss.insertSheet('Data_QuaySo');
        sheet.hideSheet();
    }
    sheet.clear();
    var chunks = [];
    for (var i = 0; i < jsonData.length; i += 45000) {
        chunks.push([jsonData.substring(i, i + 45000)]);
    }
    if (chunks.length > 0) {
        sheet.getRange(1, 1, chunks.length, 1).setValues(chunks);
    }
    return true;
}

function getQuaySoData() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Data_QuaySo');
    if (!sheet) return null;
    var values = sheet.getDataRange().getValues();
    if (!values || values.length === 0) return null;
    return values.map(function (row) { return row[0] || ''; }).join('');
}


// --- HỆ THỐNG KẾ HOẠCH ---
function saveKeHoachTitle(ca, title, currentUser) {
    incrementDataVersion();
    if (ca === "Ngày") {
        PropertiesService.getDocumentProperties().setProperty('KE_HOACH_TITLE_DAY', title);
    } else {
        PropertiesService.getDocumentProperties().setProperty('KE_HOACH_TITLE_NIGHT', title);
    }
    return { success: true };
}

function saveKeHoachDataMulti(multiDataArr, currentUser) {
    var lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (e) { return { success: false, message: "Hệ thống bận!" }; }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("KE_HOACH");
        var data = sheet.getDataRange().getValues();

        var rowsToAppend = [];
        for (var i = 0; i < multiDataArr.length; i++) {
            var newRow = multiDataArr[i];
            var ca = newRow[0];
            var maThe = newRow[1];

            var foundIdx = -1;
            for (var j = data.length - 1; j >= 1; j--) {
                if (data[j][1] == maThe && data[j][0] == ca) {
                    foundIdx = j;
                    break;
                }
            }
            if (foundIdx !== -1) {
                // Update existing row
                sheet.getRange(foundIdx + 1, 1, 1, 11).setValues([newRow]);
                data[foundIdx] = newRow;
            } else {
                rowsToAppend.push(newRow);
            }
        }

        if (rowsToAppend.length > 0) {
            sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 11).setValues(rowsToAppend);
        }

        logAction("Lưu Kế Hoạch", currentUser, "Đã lưu kế hoạch cho " + multiDataArr.length + " người");
        return { success: true, data: sheet.getDataRange().getDisplayValues() };
    } catch (e) {
        return { success: false, message: e.toString() };
    } finally {
        lock.releaseLock();
    }
}

function updateKhCell(maThe, ca, colIndex, newValue, currentUser) {
    var lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (e) { return { success: false, message: "Hệ thống bận!" }; }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("KE_HOACH");
        var data = sheet.getDataRange().getDisplayValues();
        for (var i = data.length - 1; i >= 1; i--) {
            if (data[i][1] === maThe && data[i][0] === ca) {
                if (newValue === '- Trống -') newValue = '';
                sheet.getRange(i + 1, parseInt(colIndex) + 1).setValue(newValue);
                logAction("Sửa Kế Hoạch", currentUser, "Sửa kế hoạch của: " + maThe + " (Ca " + ca + "), cột " + colIndex + " thành: " + newValue);
                return { success: true, data: sheet.getDataRange().getDisplayValues() };
            }
        }
        return { success: false, message: "Không tìm thấy dòng dữ liệu" };
    } catch (e) {
        return { success: false, message: e.toString() };
    } finally {
        lock.releaseLock();
    }
}

function clearKeHoachWeek(ca, currentUser) {
    var lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (e) { return { success: false, message: "Hệ thống bận!" }; }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("KE_HOACH");
        var data = sheet.getDataRange().getDisplayValues();
        var numRowsUpdated = 0;

        for (var i = 1; i < data.length; i++) {
            if (data[i][0] === ca) {
                // Xóa từ cột 5 đến 11 (Thứ 2 đến CN, index 1-based)
                sheet.getRange(i + 1, 5, 1, 7).clearContent();
                numRowsUpdated++;
            }
        }
        logAction("Xóa Tuần Kế Hoạch", currentUser, "Xóa toàn bộ lịch tuần Ca " + ca);
        return { success: true, data: sheet.getDataRange().getDisplayValues(), message: "Đã xóa " + numRowsUpdated + " người" };
    } catch (e) {
        return { success: false, message: e.toString() };
    } finally {
        lock.releaseLock();
    }
}

function swapKeHoachShiftsData(currentUser) {
    var lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (e) { return { success: false, message: "Hệ thống bận!" }; }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("KE_HOACH");
        var data = sheet.getDataRange().getValues();
        var numRowsUpdated = 0;

        for (var i = 1; i < data.length; i++) {
            var ca = String(data[i][0]).trim();
            if (ca === 'Ngày' || ca === 'Ngay') {
                data[i][0] = 'Đêm';
                for (var j = 4; j <= 10; j++) data[i][j] = ''; // Xóa lịch từ Thứ 2 đến CN
                numRowsUpdated++;
            } else if (ca === 'Đêm' || ca === 'Dem') {
                data[i][0] = 'Ngày';
                for (var j = 4; j <= 10; j++) data[i][j] = '';
                numRowsUpdated++;
            }
        }

        if (numRowsUpdated > 0) {
            sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
        }

        logAction("Đảo Ca Hàng Loạt", currentUser, "Đã đảo ca Ngày <-> Đêm cho " + numRowsUpdated + " người và làm trống lịch tuần.");
        return { success: true, data: sheet.getDataRange().getDisplayValues(), message: "Đã đảo ca thành công" };
    } catch (e) {
        return { success: false, message: e.toString() };
    } finally {
        lock.releaseLock();
    }
}

function saveKeHoachStartDate(ca, dateStr) {
    incrementDataVersion();
    var key = ca === 'Ngày' ? 'KE_HOACH_DATE_DAY' : 'KE_HOACH_DATE_NIGHT';
    PropertiesService.getDocumentProperties().setProperty(key, dateStr);
    return { success: true };
}

function deleteKeHoachData(maThe, ca, currentUser) {
    var lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (e) { return { success: false, message: "Hệ thống bận!" }; }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("KE_HOACH");
        var data = sheet.getDataRange().getDisplayValues();
        for (var i = data.length - 1; i >= 1; i--) {
            if (data[i][1] === maThe && data[i][0] === ca) {
                sheet.deleteRow(i + 1);
                logAction("Xóa Kế Hoạch", currentUser, "Đã xóa kế hoạch của: " + maThe + " (Ca " + ca + ")");
                return { success: true, data: sheet.getDataRange().getDisplayValues() };
            }
        }
        return { success: false, message: "Không tìm thấy dữ liệu" };
    } catch (e) {
        return { success: false, message: e.toString() };
    } finally {
        lock.releaseLock();
    }
}

// --- MARQUEE ---
function saveMarqueeSettings(text, expiry, speed, color, size) {
    incrementDataVersion();
    PropertiesService.getDocumentProperties().setProperty('MARQUEE_TEXT', text || '');
    PropertiesService.getDocumentProperties().setProperty('MARQUEE_EXPIRY', expiry || '');
    PropertiesService.getDocumentProperties().setProperty('MARQUEE_SPEED', speed || '20');
    PropertiesService.getDocumentProperties().setProperty('MARQUEE_COLOR', color || '#ff0000');
    PropertiesService.getDocumentProperties().setProperty('MARQUEE_SIZE', size || '16');
    return { success: true, text: text, expiry: expiry, speed: speed, color: color, size: size };
}

// =============================================
// QUẢN LÝ LIỆU — Lưu/Đọc dữ liệu vào sheet riêng
// =============================================
function saveQuanLyLieuData(jsonData, jsonDates) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
    } catch (e) {
        return { success: false, message: "Hệ thống đang bận, vui lòng thử lại!" };
    }
    try {
        incrementDataVersion();
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("QUAN_LY_LIEU");
        if (!sheet) {
            sheet = ss.insertSheet("QUAN_LY_LIEU");
            sheet.getRange("A1").setValue("LIEU_DATA_JSON");
            sheet.getRange("B1").setValue("LIEU_DATES_JSON");
            sheet.getRange("C1").setValue("UPDATED_AT");
        }
        sheet.getRange("A2").setValue(jsonData || "[]");
        sheet.getRange("B2").setValue(jsonDates || "[]");
        sheet.getRange("C2").setValue(new Date().toISOString());
        SpreadsheetApp.flush();
        return { success: true };
    } catch (e) {
        return { success: false, message: e.toString() };
    } finally {
        lock.releaseLock();
    }
}

function loadQuanLyLieuData() {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("QUAN_LY_LIEU");
        if (!sheet) return { success: true, data: "[]", dates: "[]" };
        var data = sheet.getRange("A2").getValue() || "[]";
        var dates = sheet.getRange("B2").getValue() || "[]";
        return { success: true, data: data, dates: dates };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

// =============================================
// BẢNG LƯƠNG — Lưu/Đọc dữ liệu vào sheet riêng
// =============================================
function saveBangLuongData(jsonNhanVien) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
    } catch (e) {
        return { success: false, message: "Hệ thống đang bận, vui lòng thử lại!" };
    }
    try {
        incrementDataVersion();
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("BANG_LUONG");
        if (!sheet) {
            sheet = ss.insertSheet("BANG_LUONG");
            sheet.getRange("A1").setValue("NHANVIEN_JSON");
            sheet.getRange("B1").setValue("UPDATED_AT");
        }
        sheet.getRange("A2").setValue(jsonNhanVien || "[]");
        sheet.getRange("B2").setValue(new Date().toISOString());
        SpreadsheetApp.flush();
        return { success: true };
    } catch (e) {
        return { success: false, message: e.toString() };
    } finally {
        lock.releaseLock();
    }
}

function loadBangLuongData() {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("BANG_LUONG");
        if (!sheet) return { success: true, data: "[]" };
        var data = sheet.getRange("A2").getValue() || "[]";
        return { success: true, data: data };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

// =============================================
// VỆ SINH & LỊCH TUẦN — Backend Functions (v2 with TUẦN)
// =============================================

function _vsSheetName(caType) { return caType === 'vs-ngay' ? 'VS_NGAY' : 'VS_DEM'; }
function _ltSheetName(caType) { return caType === 'lt-ngay' ? 'LT_NGAY' : 'LT_DEM'; }

// Migration: nếu sheet cũ không có cột TUẦN, tự chèn vào
function _migrateAddTuanCol(sheet) {
    try {
        var firstHeader = sheet.getRange(1, 1).getValue();
        if (firstHeader === 'TUẦN') return;
        sheet.insertColumnBefore(1);
        sheet.getRange(1, 1).setValue('TUẦN');
        var lastRow = sheet.getLastRow();
        if (lastRow > 1) {
            var wk = _serverWeekLabel();
            var vals = [];
            for (var i = 0; i < lastRow - 1; i++) vals.push([wk]);
            sheet.getRange(2, 1, lastRow - 1, 1).setValues(vals);
        }
    } catch (e) { Logger.log('Migration error: ' + e); }
}

function _serverWeekLabel() {
    var now = new Date();
    var day = now.getDay();
    var diff = (day === 0) ? -6 : 1 - day;
    var mon = new Date(now.getTime() + diff * 86400000);
    var sat = new Date(mon.getTime() + 5 * 86400000);
    var fmt = function (d) { return ('0' + d.getDate()).slice(-2) + '/' + ('0' + (d.getMonth() + 1)).slice(-2); };
    return fmt(mon) + '\u2013' + fmt(sat);
}

function _getLtSlotConfig() {
    var cfg = PropertiesService.getDocumentProperties().getProperty('LT_SLOT_CONFIG');
    if (cfg) try { return JSON.parse(cfg); } catch (e) { }
    return ['10:30', '02:00', '02:07'];
}

function saveLtSlotConfig(slots, currentUser) {
    incrementDataVersion();
    PropertiesService.getDocumentProperties().setProperty('LT_SLOT_CONFIG', JSON.stringify(slots));
    logAction('Sửa LT Slots', currentUser, 'Slots: ' + slots.join(', '));
    return { success: true, slots: slots };
}

function addLtSlot(newSlotName, currentUser) {
    var cfg = _getLtSlotConfig();
    cfg.push(newSlotName);
    return saveLtSlotConfig(cfg, currentUser);
}

function editLtSlot(idx, newName, currentUser) {
    var cfg = _getLtSlotConfig();
    if (idx >= 0 && idx < cfg.length) { cfg[idx] = newName; return saveLtSlotConfig(cfg, currentUser); }
    return { success: false, message: 'Invalid index' };
}

function _normalizeWeekKey(wk) {
    if (!wk) return '';
    return wk.toString().replace(/[\u2013\u2014-]/g, '-').replace(/\s+/g, '').trim();
}

// Auto-renumber STT within a week
function _renumberStt(sheet, tuanKey) {
    var data = sheet.getDataRange().getDisplayValues();
    var stt = 1;
    for (var i = 1; i < data.length; i++) {
        if (_normalizeWeekKey(data[i][0]) === _normalizeWeekKey(tuanKey)) {
            sheet.getRange(i + 1, 2).setValue(stt);
            stt++;
        }
    }
}

// --- VS CRUD (with TUẦN) ---
function saveVeSinhRow(caType, tuanKey, stt, maThe, hoTen, currentUser) {
    var lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (e) { return { success: false, message: 'Hệ thống bận!' }; }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(_vsSheetName(caType));
        if (!sheet) return { success: false, message: 'Không tìm thấy sheet!' };
        var data = sheet.getDataRange().getDisplayValues();
        for (var i = 1; i < data.length; i++) {
            if (_normalizeWeekKey(data[i][0]) === _normalizeWeekKey(tuanKey) && data[i][2] === maThe) return { success: false, message: 'Mã thẻ ' + maThe + ' đã tồn tại trong tuần này!' };
        }
        sheet.appendRow([tuanKey, stt, maThe, hoTen, '', '', '', '', '', '']);
        logAction('Thêm VS', currentUser, caType + ' tuần ' + tuanKey + ': ' + maThe);
        return { success: true, data: sheet.getDataRange().getDisplayValues() };
    } catch (e) { return { success: false, message: e.toString() }; }
    finally { lock.releaseLock(); }
}

function deleteVeSinhRow(caType, tuanKey, maThe, currentUser) {
    var lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (e) { return { success: false, message: 'Hệ thống bận!' }; }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(_vsSheetName(caType));
        var data = sheet.getDataRange().getDisplayValues();
        for (var i = data.length - 1; i >= 1; i--) {
            if (_normalizeWeekKey(data[i][0]) === _normalizeWeekKey(tuanKey) && data[i][2] === maThe) {
                sheet.deleteRow(i + 1);
                _renumberStt(sheet, tuanKey);
                logAction('Xóa VS', currentUser, caType + ': ' + maThe);
                return { success: true, data: sheet.getDataRange().getDisplayValues() };
            }
        }
        return { success: false, message: 'Không tìm thấy dòng' };
    } catch (e) { return { success: false, message: e.toString() }; }
    finally { lock.releaseLock(); }
}

function toggleVeSinhCell(caType, tuanKey, maThe, colIndex, newValue, currentUser) {
    var lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (e) { return { success: false, message: 'Hệ thống bận!' }; }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(_vsSheetName(caType));
        var data = sheet.getDataRange().getDisplayValues();
        for (var i = 1; i < data.length; i++) {
            if (_normalizeWeekKey(data[i][0]) === _normalizeWeekKey(tuanKey) && data[i][2] === maThe) {
                sheet.getRange(i + 1, parseInt(colIndex) + 1).setValue(newValue);
                return { success: true, data: sheet.getDataRange().getDisplayValues() };
            }
        }
        return { success: false, message: 'Không tìm thấy dòng' };
    } catch (e) { return { success: false, message: e.toString() }; }
    finally { lock.releaseLock(); }
}

function saveVeSinhFull(caType, tuanKey, allRows, currentUser) {
    var lock = LockService.getScriptLock();
    try { lock.waitLock(15000); } catch (e) { return { success: false, message: 'Hệ thống bận!' }; }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(_vsSheetName(caType));
        var data = sheet.getDataRange().getDisplayValues();
        // Delete rows for this week only (bottom-up)
        for (var i = data.length - 1; i >= 1; i--) {
            if (_normalizeWeekKey(data[i][0]) === _normalizeWeekKey(tuanKey)) sheet.deleteRow(i + 1);
        }
        if (allRows.length > 0) {
            sheet.getRange(sheet.getLastRow() + 1, 1, allRows.length, allRows[0].length).setValues(allRows);
        }
        logAction('Xáo VS', currentUser, caType + ' tuần ' + tuanKey);
        return { success: true, data: sheet.getDataRange().getDisplayValues() };
    } catch (e) { return { success: false, message: e.toString() }; }
    finally { lock.releaseLock(); }
}

// --- LT CRUD (with TUẦN) ---
function saveLichTuanRow(caType, tuanKey, stt, maThe, hoTen, currentUser) {
    var lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (e) { return { success: false, message: 'Hệ thống bận!' }; }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(_ltSheetName(caType));
        if (!sheet) return { success: false, message: 'Không tìm thấy sheet!' };
        var data = sheet.getDataRange().getDisplayValues();
        for (var i = 1; i < data.length; i++) {
            if (_normalizeWeekKey(data[i][0]) === _normalizeWeekKey(tuanKey) && data[i][2] === maThe) return { success: false, message: 'Mã thẻ ' + maThe + ' đã tồn tại!' };
        }
        var slots = _getLtSlotConfig();
        var row = [tuanKey, stt, maThe, hoTen];
        for (var j = 0; j < slots.length * 6; j++) row.push('');
        sheet.appendRow(row);
        logAction('Thêm LT', currentUser, caType + ': ' + maThe);
        return { success: true, data: sheet.getDataRange().getDisplayValues() };
    } catch (e) { return { success: false, message: e.toString() }; }
    finally { lock.releaseLock(); }
}

function deleteLichTuanRow(caType, tuanKey, maThe, currentUser) {
    var lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (e) { return { success: false, message: 'Hệ thống bận!' }; }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(_ltSheetName(caType));
        var data = sheet.getDataRange().getDisplayValues();
        for (var i = data.length - 1; i >= 1; i--) {
            if (_normalizeWeekKey(data[i][0]) === _normalizeWeekKey(tuanKey) && data[i][2] === maThe) {
                sheet.deleteRow(i + 1);
                _renumberStt(sheet, tuanKey);
                logAction('Xóa LT', currentUser, caType + ': ' + maThe);
                return { success: true, data: sheet.getDataRange().getDisplayValues() };
            }
        }
        return { success: false, message: 'Không tìm thấy dòng' };
    } catch (e) { return { success: false, message: e.toString() }; }
    finally { lock.releaseLock(); }
}

function toggleLichTuanCell(caType, tuanKey, maThe, colIndex, newValue, currentUser) {
    var lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (e) { return { success: false, message: 'Hệ thống bận!' }; }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(_ltSheetName(caType));
        var data = sheet.getDataRange().getDisplayValues();
        for (var i = 1; i < data.length; i++) {
            if (_normalizeWeekKey(data[i][0]) === _normalizeWeekKey(tuanKey) && data[i][2] === maThe) {
                sheet.getRange(i + 1, parseInt(colIndex) + 1).setValue(newValue);
                return { success: true, data: sheet.getDataRange().getDisplayValues() };
            }
        }
        return { success: false, message: 'Không tìm thấy dòng' };
    } catch (e) { return { success: false, message: e.toString() }; }
    finally { lock.releaseLock(); }
}

function saveLichTuanFull(caType, tuanKey, allRows, currentUser) {
    var lock = LockService.getScriptLock();
    try { lock.waitLock(15000); } catch (e) { return { success: false, message: 'Hệ thống bận!' }; }
    try {
        incrementDataVersion();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(_ltSheetName(caType));
        var data = sheet.getDataRange().getDisplayValues();
        for (var i = data.length - 1; i >= 1; i--) {
            if (_normalizeWeekKey(data[i][0]) === _normalizeWeekKey(tuanKey)) sheet.deleteRow(i + 1);
        }
        if (allRows.length > 0) {
            sheet.getRange(sheet.getLastRow() + 1, 1, allRows.length, allRows[0].length).setValues(allRows);
        }
        logAction('Xáo LT', currentUser, caType + ' tuần ' + tuanKey);
        return { success: true, data: sheet.getDataRange().getDisplayValues() };
    } catch (e) { return { success: false, message: e.toString() }; }
    finally { lock.releaseLock(); }
}

// =======================================================================
// ==================== TELEGRAM BOT CREATOR BETA MODULE =====================
// =======================================================================
// Hướng dẫn thiết lập lần đầu:
//   1. Nhắn một tin bất kỳ cho bot của bạn trên Telegram
//   2. Vào Apps Script Editor, chạy hàm: getTelegramChatId()
//   3. Xem kết quả trong Execution Log, copy chat_id
//   4. Chạy hàm: setTelegramChatId("CHAT_ID_VỪA_LẤY")
//   5. Chạy hàm: testTelegramBot() để kiểm tra
// =======================================================================

var TELEGRAM_BOT_TOKEN_DEFAULT = "8681514331:AAEGCQcS-Xm4BAssIXHsVMfXkyuxWtw_4Mc";

// Lưu Telegram Bot Token (gọi từ Apps Script Editor nếu cần đổi token)
function setTelegramChatId() {
    var chatId = "5455995122";
    var token = "8681514331:AAEGCQcS-Xm4BAssIXHsVMfXkyuxWtw_4Mc"; // Cập nhật token mới

    PropertiesService.getScriptProperties().setProperty("TELEGRAM_CHAT_ID", chatId);
    PropertiesService.getScriptProperties().setProperty("TELEGRAM_BOT_TOKEN", token);
    Logger.log("Đã lưu Telegram Chat ID: " + chatId + " và Token mới.");
}

// Cài đặt Webhook cho Telegram Bot (Chỉ cần chạy 1 lần)
function setupTelegramWebhook() {
    var props = PropertiesService.getScriptProperties();
    var token = props.getProperty("TELEGRAM_BOT_TOKEN") || TELEGRAM_BOT_TOKEN_DEFAULT;

    // NẾU BẠN CÓ ĐƯỜNG LINK WEB APP MỚI, HÃY XOÁ DÒNG DƯỚI VÀ DÁN LINK VÀO TRONG DẤU NGOẶC KÉP:
    // Ví dụ: var webAppUrl = "https://script.google.com/macros/s/AKfycbxlf.../exec";
    var webAppUrl = "https://script.google.com/macros/s/AKfycbxlFfS1PSge5rh-EVE9ZpY1QVxauPpEB31SxR03CBKtIap1cFHv7lWxZ8Owu6VV4KtV/exec";

    if (!webAppUrl || webAppUrl.indexOf("script.google.com") === -1) {
        SpreadsheetApp.getUi().alert("Lỗi: Bạn phải 'Deploy as Web App' trước khi chạy hàm này!");
        return;
    }

    // Đảm bảo URL là /exec thay vì /dev để máy chủ Telegram có thể truy cập được
    webAppUrl = webAppUrl.replace("/dev", "/exec");

    var url = "https://api.telegram.org/bot" + token + "/setWebhook?url=" + encodeURIComponent(webAppUrl) + "&drop_pending_updates=True&secret_token=CTQSYSTEM";
    try {
        var response = UrlFetchApp.fetch(url);
        var msg = "Kết quả cài đặt Webhook Telegram:\n" + response.getContentText();
        Logger.log(msg);
        SpreadsheetApp.getUi().alert(msg);
    } catch (e) {
        Logger.log("Lỗi: " + e.toString());
        SpreadsheetApp.getUi().alert("Lỗi: " + e.toString());
    }
}



// Lấy chat_id — Gọi SAU KHI đã nhắn tin cho bot trên Telegram
function getTelegramChatId() {
    var props = PropertiesService.getScriptProperties();
    var token = props.getProperty("TELEGRAM_BOT_TOKEN") || TELEGRAM_BOT_TOKEN_DEFAULT;
    var url = "https://api.telegram.org/bot" + token + "/getUpdates";
    try {
        // timeout=0: không chờ, lấy ngay; offset=-100: lấy 100 tin nhắn gần nhất
        var resp = UrlFetchApp.fetch(url + "?timeout=0&offset=-100&limit=100", { muteHttpExceptions: true });
        var rawJson = resp.getContentText();
        var result = JSON.parse(rawJson);
        Logger.log("Ket qua getUpdates: " + rawJson);

        if (result.ok && result.result && result.result.length > 0) {
            var updates = result.result;
            var chatId = null;
            for (var i = 0; i < updates.length; i++) {
                var upd = updates[i];
                // Cấu trúc Telegram: result[i].message.chat.id
                if (upd.message && upd.message.chat && upd.message.chat.id) {
                    chatId = upd.message.chat.id;
                    break;
                }
                // Fallback: result[i].message.from.id
                if (upd.message && upd.message.from && upd.message.from.id) {
                    chatId = upd.message.from.id;
                    break;
                }
                // Fallback 2: result[i].chat.id (một số event khác)
                if (upd.chat && upd.chat.id) {
                    chatId = upd.chat.id;
                    break;
                }
            }
            if (chatId) {
                Logger.log("=== TIM THAY CHAT ID: " + chatId + " ===");
                Logger.log(">>> DA TU DONG LUU TELEGRAM_CHAT_ID!");
                props.setProperty("TELEGRAM_CHAT_ID", chatId.toString());
                return chatId;
            }
        }
        Logger.log("Chua tim thay chat_id nao. Hay nhan tin cho bot tren Telegram truoc roi chay lai!");
        return null;
    } catch (e) {
        Logger.log("Loi getTelegramChatId: " + e.toString());
        return null;
    }
}

// HÀM GỬI THÔNG BÁO TELEGRAM (CORE) - Hỗ trợ nhiều người nhận
function sendTelegramNotify(message, specificChatId) {
    try {
        var props = PropertiesService.getScriptProperties();
        var token = props.getProperty("TELEGRAM_BOT_TOKEN") || TELEGRAM_BOT_TOKEN_DEFAULT;

        var chatIds = [];
        if (specificChatId) {
            chatIds.push(specificChatId.toString());
        } else {
            var chatIdsStr = props.getProperty("TELEGRAM_CHAT_ID");
            if (!chatIdsStr) return false;
            chatIds = chatIdsStr.split(",");
        }

        var url = "https://api.telegram.org/bot" + token + "/sendMessage";
        var allSuccess = true;

        for (var i = 0; i < chatIds.length; i++) {
            var id = chatIds[i].trim();
            if (!id) continue;

            var payload = { chat_id: id, text: message };
            var options = {
                method: "post",
                contentType: "application/json",
                payload: JSON.stringify(payload),
                muteHttpExceptions: true
            };
            var resp = UrlFetchApp.fetch(url, options);
            var result = JSON.parse(resp.getContentText());
            if (!result.ok) {
                Logger.log("Telegram gui that bai cho ID " + id + ": " + resp.getContentText());
                allSuccess = false;
            }
        }
        return allSuccess;
    } catch (e) {
        Logger.log("sendTelegramNotify loi: " + e.toString());
        return false;
    }
}


// TEST — Chạy từ Apps Script Editor để kiểm tra
function testTelegramBot() {
    var now = new Date().toLocaleString("vi-VN");
    var msg = "CTQ SYSTEM - KIEM TRA KET NOI\n" +
        "Bot Telegram da hoat dong!\n" +
        "Thoi gian: " + now;
    var ok = sendTelegramNotify(msg);
    Logger.log(ok ? "Gui thanh cong!" : "Gui that bai! Kiem tra log de biet id nao loi.");
    return ok;
}

// ==========================================
// --- ĐIỂM DANH / ATTENDANCE ---
// ==========================================

function getDiemDanhEmployees() {
    try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("出勤数据");
        if (!sheet) return { success: false, message: "Không tìm thấy sheet 出勤数据" };
        var data = sheet.getDataRange().getDisplayValues();
        return { success: true, data: data };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function saveDiemDanh(sheetName, headerRow, dataRows) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
    } catch (e) {
        return { success: false, message: "Hệ thống đang bận, vui lòng thử lại!" };
    }
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
            sheet = ss.insertSheet(sheetName);
        } else {
            sheet.clear();
            sheet.clearFormats();
        }
        var maxCols = headerRow.length;

        // ROW 1: Title row - merged
        var titleText = "Bảng theo dõi chuyên cần CTQ Tháng / CTQ 出勤追踪表";
        // Tự trích tháng từ sheetName (VD: ĐIỂM_DANH_T7_2026 → Tháng 7)
        var mMatch = sheetName.match(/T(\d+)/);
        if (mMatch) titleText = "Bảng theo dõi chuyên cần CTQ Tháng " + mMatch[1] + "/ CTQ 出勤追踪表" + mMatch[1] + "月";
        sheet.getRange(1, 1).setValue(titleText);
        sheet.getRange(1, 1, 1, maxCols).merge()
            .setFontWeight("bold").setFontSize(21).setHorizontalAlignment("center")
            .setFontColor("#1a5632").setBackground("#ffffff");
        sheet.setRowHeight(1, 45);

        // ROW 2: Header row
        sheet.getRange(2, 1, 1, maxCols).setValues([headerRow]);
        sheet.getRange(2, 1, 1, maxCols)
            .setFontWeight("bold").setFontSize(21).setFontFamily("Times New Roman")
            .setBackground("#1a5632").setFontColor("#ffffff")
            .setHorizontalAlignment("center").setVerticalAlignment("middle")
            .setWrap(true);
        sheet.setRowHeight(2, 50);

        // Summary columns (last 3) - different header colors
        if (maxCols >= 3) {
            sheet.getRange(2, maxCols - 2).setBackground("#0d47a1"); // Số ngày làm
            sheet.getRange(2, maxCols - 1).setBackground("#b71c1c"); // Số ngày nghỉ
            sheet.getRange(2, maxCols).setBackground("#e65100");     // Tỷ lệ
        }

        // ROW 3+: Data rows
        if (dataRows && dataRows.length > 0) {
            var normalizedRows = dataRows.map(function (r) {
                while (r.length < maxCols) r.push("");
                return r.slice(0, maxCols);
            });
            sheet.getRange(3, 1, normalizedRows.length, maxCols).setValues(normalizedRows);
            sheet.getRange(3, 1, normalizedRows.length, maxCols)
                .setFontWeight("bold").setFontSize(17).setFontFamily("Times New Roman")
                .setHorizontalAlignment("center").setVerticalAlignment("middle");

            // Zebra striping & yellow background for '周日'
            for (var i = 0; i < normalizedRows.length; i++) {
                if (i % 2 === 1) {
                    sheet.getRange(3 + i, 1, 1, maxCols).setBackground("#f1f8e9");
                }
                for (var c = 0; c < maxCols; c++) {
                    if (normalizedRows[i][c] === '周日') {
                        sheet.getRange(3 + i, c + 1).setBackground("#fff59d").setFontColor("#000000");
                    }
                }
                sheet.setRowHeight(3 + i, 45); // Increased row height for font size 17
            }
        }

        // Borders on all data area
        var totalRows = 2 + (dataRows ? dataRows.length : 0);
        sheet.getRange(2, 1, totalRows - 1, maxCols).setBorder(true, true, true, true, true, true,
            "#333333", SpreadsheetApp.BorderStyle.SOLID);

        // Column widths (no joinDate column now: STT, ID, NAME, dates..., summary)
        sheet.setColumnWidth(1, 60);  // STT
        sheet.setColumnWidth(2, 160); // ID
        sheet.setColumnWidth(3, 260); // NAME
        // Date columns
        for (var c = 4; c <= maxCols - 3; c++) {
            sheet.setColumnWidth(c, 60);
        }
        // Summary columns
        if (maxCols >= 3) {
            sheet.setColumnWidth(maxCols - 2, 80);
            sheet.setColumnWidth(maxCols - 1, 80);
            sheet.setColumnWidth(maxCols, 80);
        }

        // Freeze header
        sheet.setFrozenRows(2);

        SpreadsheetApp.flush();
        return { success: true };
    } catch (e) {
        return { success: false, message: e.toString() };
    } finally {
        lock.releaseLock();
    }
}

function loadDiemDanh(sheetName) {
    try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
        if (!sheet) return { success: false, message: "Không tìm thấy sheet: " + sheetName };
        var data = sheet.getDataRange().getDisplayValues();
        return { success: true, data: data };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function getListDiemDanhSheets() {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheets = ss.getSheets();
        var result = [];
        for (var i = 0; i < sheets.length; i++) {
            var name = sheets[i].getName();
            if (name.indexOf("ĐIỂM_DANH_") === 0) {
                result.push(name);
            }
        }
        return { success: true, data: result };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

// ==========================================
// --- LỖI CTQ / CTQ ERROR TRACKING ---
// ==========================================

function saveLoiCTQ(sheetName, headerRow, dataRows, notesRows) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
    } catch (e) {
        return { success: false, message: "Hệ thống đang bận, vui lòng thử lại!" };
    }
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
            sheet = ss.insertSheet(sheetName);
        } else {
            sheet.clear();
            sheet.clearFormats();
        }
        var maxCols = headerRow.length;

        // ROW 1: Title row - merged
        var titleText = "Bắt lỗi Điểm Mù CTQ Tháng / CTQ 盲点错误追踪表";
        var mMatch = sheetName.match(/T(\d+)/);
        if (mMatch) titleText = "Bắt lỗi Điểm Mù CTQ Tháng " + mMatch[1] + "/ CTQ 盲点错误追踪表" + mMatch[1] + "月";
        sheet.getRange(1, 1).setValue(titleText);
        sheet.getRange(1, 1, 1, maxCols).merge()
            .setFontWeight("bold").setFontSize(21).setHorizontalAlignment("center")
            .setFontColor("#b71c1c").setBackground("#ffffff");
        sheet.setRowHeight(1, 45);

        // ROW 2: Header row
        sheet.getRange(2, 1, 1, maxCols).setValues([headerRow]);
        sheet.getRange(2, 1, 1, maxCols)
            .setFontWeight("bold").setFontSize(21).setFontFamily("Times New Roman")
            .setBackground("#b71c1c").setFontColor("#ffffff")
            .setHorizontalAlignment("center").setVerticalAlignment("middle")
            .setWrap(true);
        sheet.setRowHeight(2, 50);

        // Summary columns (last 1) - TOTAL
        if (maxCols >= 1) {
            sheet.getRange(2, maxCols).setBackground("#0d47a1");
        }

        // ROW 3+: Data rows
        if (dataRows && dataRows.length > 0) {
            var normalizedRows = dataRows.map(function (r) {
                while (r.length < maxCols) r.push("");
                return r.slice(0, maxCols);
            });
            sheet.getRange(3, 1, normalizedRows.length, maxCols).setValues(normalizedRows);
            if (notesRows && notesRows.length > 0) {
                var normalizedNotes = notesRows.map(function (r) {
                    while (r.length < maxCols) r.push("");
                    return r.slice(0, maxCols);
                });
                sheet.getRange(3, 1, normalizedNotes.length, maxCols).setNotes(normalizedNotes);
            }
            sheet.getRange(3, 1, normalizedRows.length, maxCols)
                .setFontWeight("bold").setFontSize(17).setFontFamily("Times New Roman")
                .setHorizontalAlignment("center").setVerticalAlignment("middle");

            // Zebra striping & yellow background for '周日'
            for (var i = 0; i < normalizedRows.length; i++) {
                if (i % 2 === 1) {
                    sheet.getRange(3 + i, 1, 1, maxCols).setBackground("#fce4ec");
                }
                for (var c = 0; c < maxCols; c++) {
                    if (normalizedRows[i][c] === '周日') {
                        sheet.getRange(3 + i, c + 1).setBackground("#fff59d").setFontColor("#000000");
                    }
                }
                sheet.setRowHeight(3 + i, 45);
            }
        }

        // Borders
        var totalRows = 2 + (dataRows ? dataRows.length : 0);
        sheet.getRange(2, 1, totalRows - 1, maxCols).setBorder(true, true, true, true, true, true,
            "#333333", SpreadsheetApp.BorderStyle.SOLID);

        // Column widths
        sheet.setColumnWidth(1, 60);  // STT
        sheet.setColumnWidth(2, 160); // ID
        sheet.setColumnWidth(3, 120); // VỊ TRÍ
        sheet.setColumnWidth(4, 260); // NAME
        for (var c = 5; c <= maxCols - 1; c++) {
            sheet.setColumnWidth(c, 60);
        }
        if (maxCols >= 1) {
            sheet.setColumnWidth(maxCols, 80); // TOTAL
        }

        sheet.setFrozenRows(2);
        SpreadsheetApp.flush();
        return { success: true };
    } catch (e) {
        return { success: false, message: e.toString() };
    } finally {
        lock.releaseLock();
    }
}

function loadLoiCTQ(sheetName) {
    try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
        if (!sheet) return { success: false, message: "Không tìm thấy sheet: " + sheetName };
        var range = sheet.getDataRange();
        var data = range.getDisplayValues();
        var notes = range.getNotes();
        return { success: true, data: data, notes: notes };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function getListLoiCTQSheets() {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheets = ss.getSheets();
        var result = [];
        for (var i = 0; i < sheets.length; i++) {
            var name = sheets[i].getName();
            if (name.indexOf("LOI_CTQ_") === 0) {
                result.push(name);
            }
        }
        return { success: true, data: result };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

// ==========================================
// --- BÁO CHUYÊN CẦN (ATTENDANCE REPORT) ---
// ==========================================

function saveBaoChuyenCan(jsonData) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var parsed = JSON.parse(jsonData);



        // 2. Vẽ bảng báo cáo trực quan lên sheet
        var visualSheetName = "Báo Chuyên Cần";
        var sheet = ss.getSheetByName(visualSheetName);
        if (!sheet) {
            sheet = ss.insertSheet(visualSheetName);
        }

        // Luôn luôn ghi đè lên bảng ở trên cùng (dòng 1), không tạo thêm bảng mới
        var startRow = 1;

        drawVisualTable(sheet, startRow, parsed.date, parsed.day, parsed.night);

        return { success: true };
    } catch (e) {
        return { success: false, message: e.toString() };
    } finally {
        lock.releaseLock();
    }
}

function loadBaoChuyenCan() {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("Báo Chuyên Cần");
        if (!sheet) return { success: true, data: null };

        var vals = sheet.getRange(1, 1, 10, 11).getValues();

        var dateCell = vals[0][0].toString();
        var dateParts = dateCell.split('\n');
        var parsedDate = dateParts.length >= 3 ? dateParts[2] : dateParts[0];

        function parseData(offset) {
            var l1 = vals[2][offset + 4] ? vals[2][offset + 4].toString() : "";
            var l2 = vals[3][offset + 4] ? vals[3][offset + 4].toString() : "";
            var l3 = vals[4][offset + 4] ? vals[4][offset + 4].toString() : "";
            if (l1.indexOf("Line 1: ") === 0) l1 = l1.substring(8);
            if (l2.indexOf("Line 2: ") === 0) l2 = l2.substring(8);
            if (l3.indexOf("Line 3: ") === 0) l3 = l3.substring(8);

            return {
                slTong: vals[1][offset + 1],
                nghiViecRieng: vals[2][offset + 1],
                nghiTuDo: vals[3][offset + 1],
                nghiOm: vals[4][offset + 1],
                nghiPhep: vals[5][offset + 1],
                vaoMuon: vals[6][offset + 1],
                nghiViec: vals[7][offset + 1],
                tongCoMat: vals[8][offset + 1],
                hoTroMain: vals[2][offset + 3],
                hoTroSub: vals[5][offset + 3],
                ctqGiaoQuan: vals[6][offset + 3],
                traNguoiMoi: vals[8][offset + 3],
                dangDaoTao: vals[9][offset + 3],
                line1: l1,
                line2: l2,
                line3: l3,
                note1: vals[5][offset + 4],
                note2: vals[6][offset + 4],
                note3: vals[7][offset + 4],
                note4: vals[8][offset + 4],
                note5: vals[9][offset + 4]
            };
        }

        var payload = {
            date: parsedDate,
            day: parseData(0),
            night: parseData(6)
        };

        return { success: true, data: JSON.stringify(payload) };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function drawVisualTable(sheet, r, dateStr, dData, nData) {
    sheet.getRange(r, 1, 10, 11).breakApart().clearFormat().clearContent();

    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 80);
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(4, 80);
    sheet.setColumnWidth(5, 200);
    sheet.setColumnWidth(6, 20); // Divider
    sheet.setColumnWidth(7, 150);
    sheet.setColumnWidth(8, 80);
    sheet.setColumnWidth(9, 150);
    sheet.setColumnWidth(10, 80);
    sheet.setColumnWidth(11, 200);

    var vals = [];
    var bgs = [];
    for (var i = 0; i < 10; i++) {
        var rowVals = [];
        var rowBgs = [];
        for (var j = 0; j < 11; j++) { rowVals.push(""); rowBgs.push("#ffffff"); }
        vals.push(rowVals);
        bgs.push(rowBgs);
    }

    function fill(offset, sData, vName, cName) {
        vals[0][offset] = vName + "\n" + cName + "\n" + dateStr; bgs[0][offset] = "#ffeb00";
        vals[0][offset + 1] = "Báo Cáo Chuyên Cần CTQ\nCTQ 考勤报告"; bgs[0][offset + 1] = "#ffeb00"; bgs[0][offset + 2] = "#ffeb00"; bgs[0][offset + 3] = "#ffeb00"; bgs[0][offset + 4] = "#ffeb00";

        vals[1][offset] = "SL Tổng\n总量"; bgs[1][offset] = "#8bc34a"; vals[1][offset + 1] = sData.slTong || "";
        vals[1][offset + 2] = "Bộ Phận Hỗ Trợ\n支援部门"; bgs[1][offset + 2] = "#90caf9";
        vals[1][offset + 3] = "SL hỗ trợ\n支援数量"; bgs[1][offset + 3] = "#90caf9";
        vals[1][offset + 4] = "Ghi chú\n备注"; bgs[1][offset + 4] = "#90caf9";

        vals[2][offset] = "Nghỉ việc riêng\n事假"; bgs[2][offset] = "#8bc34a"; vals[2][offset + 1] = sData.nghiViecRieng || "";
        vals[2][offset + 2] = "Hỗ trợ MAIN\nMAIN 线支援"; bgs[2][offset + 2] = "#bbdefb";
        vals[2][offset + 3] = sData.hoTroMain || ""; vals[2][offset + 4] = "Line 1: " + (sData.line1 || "");

        vals[3][offset] = "Nghỉ tự do\n旷工"; bgs[3][offset] = "#8bc34a"; vals[3][offset + 1] = sData.nghiTuDo || "";
        bgs[3][offset + 2] = "#bbdefb"; vals[3][offset + 4] = "Line 2: " + (sData.line2 || "");

        vals[4][offset] = "Nghỉ ốm\n病假"; bgs[4][offset] = "#8bc34a"; vals[4][offset + 1] = sData.nghiOm || "";
        bgs[4][offset + 2] = "#bbdefb"; vals[4][offset + 4] = "Line 3: " + (sData.line3 || "");

        vals[5][offset] = "Nghỉ phép\n年假"; bgs[5][offset] = "#8bc34a"; vals[5][offset + 1] = sData.nghiPhep || "";
        vals[5][offset + 2] = "Hỗ trợ SUB\nSUB 线支援"; bgs[5][offset + 2] = "#bbdefb";
        vals[5][offset + 3] = sData.hoTroSub || ""; vals[5][offset + 4] = sData.note1 || "";

        vals[6][offset] = "Vào muộn\n迟到"; bgs[6][offset] = "#8bc34a"; vals[6][offset + 1] = sData.vaoMuon || "";
        vals[6][offset + 2] = "CTQ教官"; bgs[6][offset + 2] = "#bbdefb";
        vals[6][offset + 3] = sData.ctqGiaoQuan || ""; vals[6][offset + 4] = sData.note2 || "";

        vals[7][offset] = "Nghỉ việc\n离职"; bgs[7][offset] = "#8bc34a"; vals[7][offset + 1] = sData.nghiViec || "";
        bgs[7][offset + 2] = "#bbdefb"; vals[7][offset + 4] = sData.note3 || "";

        vals[8][offset] = "Tổng SL có mặt\n总出席人数"; bgs[8][offset] = "#ffeb00"; vals[8][offset + 1] = sData.tongCoMat || "";
        vals[8][offset + 2] = "Trả người mới\n推给新人培训组"; bgs[8][offset + 2] = "#bbdefb";
        vals[8][offset + 3] = sData.traNguoiMoi || ""; vals[8][offset + 4] = sData.note4 || "";

        vals[9][offset] = ""; bgs[9][offset] = "#ffeb00";
        vals[9][offset + 1] = "";
        vals[9][offset + 2] = "Đang Đào Tạo\n培训中"; bgs[9][offset + 2] = "#bbdefb";
        vals[9][offset + 3] = sData.dangDaoTao || ""; vals[9][offset + 4] = sData.note5 || "";
    }

    fill(0, dData || {}, "Ngày", "白班");
    fill(6, nData || {}, "ĐÊM", "夜班");

    var range = sheet.getRange(r, 1, 10, 11);
    range.setValues(vals);
    range.setBackgrounds(bgs);
    range.setFontWeight("bold");
    range.setHorizontalAlignment("center");
    range.setVerticalAlignment("middle");
    range.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    sheet.getRange(r, 6, 10, 1).setBorder(false, false, false, false, false, false);
    sheet.getRange(r, 6, 10, 1).setBackground("#ffffff");

    function mrg(rOff, cOff, numRows, numCols) { sheet.getRange(r + rOff, cOff, numRows, numCols).merge(); }

    // Day merges
    mrg(0, 2, 1, 4);
    mrg(2, 3, 3, 1);
    mrg(6, 3, 2, 1);
    mrg(8, 1, 2, 1);
    mrg(8, 2, 2, 1);

    // Night merges
    mrg(0, 8, 1, 4);
    mrg(2, 9, 3, 1);
    mrg(6, 9, 2, 1);
    mrg(8, 7, 2, 1);
    mrg(8, 8, 2, 1);

    // Values color
    sheet.getRange(r + 8, 2, 2, 1).setFontColor("red");
    sheet.getRange(r + 8, 8, 2, 1).setFontColor("red");
}



// ==========================================
// CHỨC NĂNG THEO DÕI CÔNG NHÂN (CHẤM CÔNG)
// ==========================================
function luuTheoDoiCongNhan(month, year, data, comments, colors, textColors) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000); // Đợi 10 giây

        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheetName = "CHAM_CONG_T" + month + "_" + year;
        var sheet = ss.getSheetByName(sheetName);

        if (!sheet) {
            sheet = ss.insertSheet(sheetName);
        } else {
            sheet.clear();
        }

        if (data && data.length > 0) {
            var numRows = data.length;
            var numCols = data[0].length;

            // Đảm bảo sheet có đủ kích thước
            if (sheet.getMaxRows() < numRows) {
                sheet.insertRowsAfter(sheet.getMaxRows(), numRows - sheet.getMaxRows());
            }
            if (sheet.getMaxColumns() < numCols) {
                sheet.insertColumnsAfter(sheet.getMaxColumns(), numCols - sheet.getMaxColumns());
            }

            var fullRange = sheet.getRange(1, 1, numRows, numCols);
            fullRange.setValues(data);

            // Format font size and alignment
            fullRange.setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");

            // Format borders (black solid)
            fullRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

            // Xử lý background colors theo mảng 2D để tối ưu tốc độ
            var bgColors = [];
            var notes = [];
            var fontColorsList = [];
            for (var i = 0; i < numRows; i++) {
                var rowBgs = [];
                var rowNotes = [];
                var rowFonts = [];
                for (var j = 0; j < numCols; j++) {
                    rowNotes.push(""); // Khởi tạo mảng ghi chú trống
                    rowFonts.push("#000000"); // Mặc định chữ đen
                    if (i === 0) {
                        rowBgs.push("#f1f5f9"); // Header
                    } else {
                        // Default colors
                        if (j >= 1 && j <= 5) rowBgs.push("#fef08a"); // Model -> Station
                        else if (j === 6) rowBgs.push("#bbf7d0"); // CA
                        else if (j === 7) rowBgs.push("#fef08a"); // LINE
                        else rowBgs.push(null);
                    }
                }
                bgColors.push(rowBgs);
                notes.push(rowNotes);
                fontColorsList.push(rowFonts);
            }

            // Xử lý màu sắc (Cam chủ nhật, quick fill, v.v.)
            if (colors && typeof colors === 'object') {
                for (var key in colors) {
                    var parts = key.split('_');
                    var r = parseInt(parts[0]);
                    var c = parseInt(parts[1]);
                    if (r < numRows - 1 && c < numCols) {
                        bgColors[r + 1][c] = colors[key];
                    }
                }
            }

            // Xử lý ghi chú
            if (comments && typeof comments === 'object') {
                for (var key in comments) {
                    var parts = key.split('_');
                    var r = parseInt(parts[0]);
                    var c = parseInt(parts[1]);
                    if (r < numRows - 1 && c < numCols) {
                        notes[r + 1][c] = comments[key];
                        bgColors[r + 1][c] = "#fee2e2"; // Nền đỏ nhạt cho ô có ghi chú
                    }
                }
            }

            // Xử lý màu sắc chữ
            if (textColors && typeof textColors === 'object') {
                for (var key in textColors) {
                    var parts = key.split('_');
                    var r = parseInt(parts[0]);
                    var c = parseInt(parts[1]);
                    if (r < numRows - 1 && c < numCols) {
                        fontColorsList[r + 1][c] = textColors[key];
                    }
                }
            }

            // Đổ toàn bộ màu, ghi chú và màu chữ
            fullRange.setBackgrounds(bgColors);
            fullRange.setNotes(notes);
            fullRange.setFontColors(fontColorsList);

            // Cột CA đổi màu chữ
            if (numRows > 1) {
                sheet.getRange(2, 7, numRows - 1, 1).setFontColor("#166534");
            }

            // Resize columns for better visibility with font size 16
            sheet.setColumnWidth(1, 60); // STT
            sheet.setColumnWidth(2, 100); // Model
            sheet.setColumnWidth(3, 120); // Area
            sheet.setColumnWidth(4, 150); // ID
            sheet.setColumnWidth(5, 250); // Name
            sheet.setColumnWidth(6, 300); // Station
            sheet.setColumnWidth(7, 80);  // CA
            sheet.setColumnWidth(8, 120); // Line

            // Cố định dòng 1 và 8 cột đầu
            sheet.setFrozenRows(1);
            sheet.setFrozenColumns(8);
        }

        return { success: true, sheetName: sheetName };
    } catch (e) {
        return { success: false, message: e.toString() };
    } finally {
        lock.releaseLock();
    }
}

function taiTheoDoiCongNhan(month, year) {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheetName = "CHAM_CONG_T" + month + "_" + year;
        var sheet = ss.getSheetByName(sheetName);

        if (!sheet) {
            return { success: false, message: "Không tìm thấy dữ liệu tháng " + month + "/" + year };
        }

        var fullRange = sheet.getDataRange();
        var fullValues = fullRange.getDisplayValues();
        if (fullValues.length < 1) {
            return { success: false, message: "Bảng dữ liệu trống!" };
        }

        var bgColors = fullRange.getBackgrounds();
        var fontColors = fullRange.getFontColors();
        var notes = fullRange.getNotes();

        var data = fullValues.slice(1);
        var header = fullValues[0];

        // Trích xuất tdcnDates (từ cột 8 đến cuối trừ 14 cột tính toán)
        var tdcnDates = [];
        if (header.length > 8 + 14) {
            tdcnDates = header.slice(8, header.length - 14);
        }

        // Trích xuất comments và colors thành object flat
        var comments = {};
        var colors = {};
        var textColors = {};

        for (var i = 1; i < fullValues.length; i++) {
            for (var j = 0; j < header.length; j++) {
                var cNote = notes[i][j];
                if (cNote) comments[(i - 1) + "_" + j] = cNote;

                var cBg = bgColors[i][j];
                // Không lưu các màu mặc định để giảm băng thông
                if (cBg && cBg !== "#ffffff" && cBg !== "#fef08a" && cBg !== "#bbf7d0" && cBg !== "#f1f5f9" && cBg !== "#fee2e2" && cBg !== "#e5e7eb") {
                    colors[(i - 1) + "_" + j] = cBg;
                }

                var cFont = fontColors[i][j];
                if (cFont && cFont !== "#000000") {
                    textColors[(i - 1) + "_" + j] = cFont;
                }
            }
        }

        return {
            success: true,
            data: data,
            dates: tdcnDates,
            comments: comments,
            colors: colors,
            textColors: textColors
        };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

// Lấy dữ liệu Điểm Mù từ sheet DATA
function getDiemMuTheoThang(month, year) {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("DATA");
        if (!sheet) return {};

        var data = sheet.getDataRange().getValues();
        var result = {};

        var prevM = month - 1;
        var prevY = year;
        if (prevM === 0) { prevM = 12; prevY = year - 1; }

        var startDate = new Date(prevY, prevM - 1, 26);
        var endDate = new Date(year, month - 1, 25);
        startDate.setHours(0, 0, 0, 0);
        endDate.setHours(23, 59, 59, 999);

        for (var i = 1; i < data.length; i++) {
            var row = data[i];
            var dateVal = row[11]; // Cột L (index 11) - Ngày
            var empId = row[1]; // Cột B (index 1) - Mã thẻ
            var colF = row[5]; // Cột F (index 5) - Số lỗi

            if (dateVal) {
                var dTime = null;
                if (typeof dateVal.getTime === 'function') {
                    dTime = dateVal.getTime();
                } else if (typeof dateVal === 'string') {
                    var p = dateVal.split('/');
                    if (p.length === 3) {
                        dTime = new Date(p[2], p[1] - 1, p[0]).getTime();
                    }
                }

                if (dTime !== null) {
                    if (dTime >= startDate.getTime() && dTime <= endDate.getTime()) {
                        if (empId) {
                            empId = empId.toString().trim();
                            if (empId !== "") {
                                var num = parseFloat(colF);
                                if (!isNaN(num) && num >= 1) {
                                    result[empId] = (result[empId] || 0) + num;
                                }
                            }
                        }
                    }
                }
            }
        }
        return result;
    } catch (e) {
        return {};
    }
}

// ==========================================
// CÀI ĐẶT BẢNG ĐIỀN NHANH (QUICK FILL CTQ)
// ==========================================
function getQuickFillData() {
    try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ADMIN_DATA");
        if (!sheet) return null;

        // Cột L (12), M (13), N (14), O (15) - O là màu chữ
        var lastRow = sheet.getLastRow();
        if (lastRow < 2) return null;

        var dataRange = sheet.getRange(2, 12, lastRow - 1, 4).getValues();
        var result = [];
        for (var i = 0; i < dataRange.length; i++) {
            var v = (dataRange[i][0] || "").toString().trim();
            var d = (dataRange[i][1] || "").toString().trim();
            var c = (dataRange[i][2] || "").toString().trim();
            var tc = (dataRange[i][3] || "").toString().trim();
            if (v !== "") {
                var item = { v: v, d: d, c: c };
                if (tc) item.tc = tc;
                result.push(item);
            }
        }
        return result.length > 0 ? result : null;
    } catch (e) {
        return null;
    }
}

function saveQuickFillData(dataArray) {
    try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ADMIN_DATA");
        if (!sheet) return "Lỗi: Không tìm thấy sheet ADMIN_DATA";

        // Xóa dữ liệu cũ ở cột L, M, N, O (từ dòng 2 trở đi)
        var lastRow = sheet.getLastRow();
        if (lastRow >= 2) {
            sheet.getRange(2, 12, lastRow, 4).clearContent();
        }

        // Thêm tiêu đề nếu chưa có
        if (lastRow < 1) {
            sheet.getRange("L1").setValue("Ký Hiệu");
            sheet.getRange("M1").setValue("Miêu Tả");
            sheet.getRange("N1").setValue("Màu Nền");
            sheet.getRange("O1").setValue("Màu Chữ");
        } else {
            if (!sheet.getRange("L1").getValue()) sheet.getRange("L1").setValue("Ký Hiệu");
            if (!sheet.getRange("M1").getValue()) sheet.getRange("M1").setValue("Miêu Tả");
            if (!sheet.getRange("N1").getValue()) sheet.getRange("N1").setValue("Màu Nền");
            if (!sheet.getRange("O1").getValue()) sheet.getRange("O1").setValue("Màu Chữ");
        }

        if (dataArray && dataArray.length > 0) {
            var output = [];
            for (var i = 0; i < dataArray.length; i++) {
                output.push([dataArray[i].v, dataArray[i].d, dataArray[i].c, dataArray[i].tc || ""]);
            }
            sheet.getRange(2, 12, output.length, 4).setValues(output);
        }

        return "SUCCESS";
    } catch (e) {
        return "Lỗi: " + e.toString();
    }
}

// Lấy danh sách các tháng đã có sheet CHAM_CONG_T
function getAvailableTdcnMonths() {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheets = ss.getSheets();
        var available = [];
        for (var i = 0; i < sheets.length; i++) {
            var name = sheets[i].getName();
            if (name.indexOf("CHAM_CONG_T") === 0) {
                var parts = name.substring(11).split("_");
                if (parts.length === 2) {
                    var m = parseInt(parts[0]);
                    var y = parseInt(parts[1]);
                    if (!isNaN(m) && !isNaN(y)) {
                        available.push({ month: m, year: y });
                    }
                }
            }
        }
        available.sort(function (a, b) {
            if (a.year !== b.year) return b.year - a.year;
            return b.month - a.month;
        });
        return available;
    } catch (e) {
        return [];
    }
}

// ==========================================
// TRA CỨU MÃ THẺ - CARD CODE LOOKUP
// ==========================================
function searchMaTheData(maTheList) {
    try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("出勤数据");
        if (!sheet) return { success: false, message: "Không tìm thấy sheet 出勤数据" };
        var data = sheet.getDataRange().getDisplayValues();
        if (data.length < 2) return { success: true, results: {} };

        // Chuẩn hóa mã thẻ input
        var searchMap = {};
        for (var i = 0; i < maTheList.length; i++) {
            var mt = String(maTheList[i]).trim().toUpperCase();
            if (mt) searchMap[mt] = [];
        }

        // Duyệt từ dòng 2 (bỏ header)
        // Col A(0)=Bộ Phận, B(1)=Mã Thẻ, C(2)=Họ Tên, D(3)=Ngày Vào, E(4)=Ngày DL, F(5)=Chấm Công, G(6)=Giờ Về, H(7)=Ca
        for (var r = 1; r < data.length; r++) {
            var maThe = String(data[r][1]).trim().toUpperCase();
            if (maThe && searchMap.hasOwnProperty(maThe)) {
                searchMap[maThe].push({
                    boPhan: data[r][0] || '',
                    hoTen: data[r][2] || '',
                    ngayVao: data[r][3] || '',
                    ngayDL: data[r][4] || '',
                    chamCong: data[r][5] || '',
                    gioVe: data[r][6] || '',
                    ca: data[r][7] || ''
                });
            }
        }

        return { success: true, results: searchMap };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function searchHoTenData(hoTenList) {
    try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("出勤数据");
        if (!sheet) return { success: false, message: "Không tìm thấy sheet 出勤数据" };
        var data = sheet.getDataRange().getDisplayValues();
        if (data.length < 2) return { success: true, results: {} };

        // Chuẩn hóa họ tên input
        var searchMap = {};
        for (var i = 0; i < hoTenList.length; i++) {
            var ht = String(hoTenList[i]).trim().toUpperCase();
            if (ht) searchMap[ht] = [];
        }

        // Duyệt từ dòng 2 (bỏ header)
        // Col A(0)=Bộ Phận, B(1)=Mã Thẻ, C(2)=Họ Tên, D(3)=Ngày Vào, E(4)=Ngày DL, F(5)=Chấm Công, G(6)=Giờ Về, H(7)=Ca
        for (var r = 1; r < data.length; r++) {
            var hoTen = String(data[r][2]).trim().toUpperCase();
            if (hoTen && searchMap.hasOwnProperty(hoTen)) {
                searchMap[hoTen].push({
                    maThe: data[r][1] || '',
                    boPhan: data[r][0] || '',
                    ngayVao: data[r][3] || '',
                    ngayDL: data[r][4] || '',
                    chamCong: data[r][5] || '',
                    gioVe: data[r][6] || '',
                    ca: data[r][7] || ''
                });
            }
        }

        return { success: true, results: searchMap };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function getAllTraCuuData() {
    try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("出勤数据");
        if (!sheet) return { success: false, message: "Không tìm thấy sheet 出勤数据" };
        var data = sheet.getDataRange().getDisplayValues();

        var optimized = [];
        for (var r = 1; r < data.length; r++) {
            optimized.push([
                data[r][0] || '', // Bộ Phận
                data[r][1] || '', // Mã Thẻ
                data[r][2] || '', // Họ Tên
                data[r][3] || '', // Ngày Vào
                data[r][4] || '', // Ngày DL
                data[r][5] || '', // Chấm Công
                data[r][6] || '', // Giờ Về
                data[r][7] || ''  // Ca
            ]);
        }
        return { success: true, data: optimized };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

// ==========================================
// QUẢN LÝ TIẾN ĐỘ CTQ
// ==========================================
function loadTienDoCTQData() {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("TIEN_DO_CTQ");
        if (!sheet) {
            return JSON.stringify({ stations: [], people: [], checks: {} });
        }
        var data = sheet.getRange("A1").getValue();
        return data ? data.toString() : JSON.stringify({ stations: [], people: [], checks: {} });
    } catch (e) {
        return null;
    }
}

function saveTienDoCTQData(jsonData) {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("TIEN_DO_CTQ");
        if (!sheet) {
            sheet = ss.insertSheet("TIEN_DO_CTQ");
            sheet.hideSheet();
        }
        sheet.getRange("A1").setValue(jsonData);
        return true;
    } catch (e) {
        throw new Error(e.toString());
    }
}

// ==========================================
// LIÊN HỆ NHÓM CTQ
// ==========================================

function getLienHeData() {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("LIEN_HE");
        if (!sheet) {
            return [];
        }
        var data = sheet.getDataRange().getValues();
        if (data.length <= 1) return [];

        var result = [];
        for (var i = 1; i < data.length; i++) {
            var row = data[i];
            if (!row[0]) continue;
            result.push({
                id: row[0],
                ngay: row[1] ? Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : "",
                maThe: row[2],
                hoTen: row[3],
                boPhan: row[4],
                cauHoi: row[5],
                giaoVien: row[6] || "",
                traLoi: row[7] || "",
                loaiHinh: row[8] || "",
                daXem: row[9] === "" ? "" : row[9],
                username: row[10] || ""
            });
        }
        return result.reverse();
    } catch (e) {
        return [];
    }
}

function addCauHoiLienHe(dataObj) {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("LIEN_HE");
        if (!sheet) return { success: false, message: "Không tìm thấy sheet LIEN_HE" };

        var newId = "CH" + new Date().getTime();
        var date = new Date();
        sheet.appendRow([newId, date, dataObj.maThe, dataObj.hoTen, dataObj.boPhan, dataObj.cauHoi, "", "", dataObj.loaiHinh || "", "", dataObj.username || ""]);

        return { success: true };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function answerCauHoiLienHe(id, gv, answer) {
    try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LIEN_HE");
        if (!sheet) return { success: false, message: "Không tìm thấy sheet LIEN_HE" };

        var data = sheet.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
            if (data[i][0] == id) {
                sheet.getRange(i + 1, 7).setValue(gv);
                sheet.getRange(i + 1, 8).setValue(answer);
                sheet.getRange(i + 1, 10).setValue(false); // Chưa xem
                var maThe = String(data[i][2]).trim().toUpperCase();
                var userSubmit = String(data[i][10] || "").trim().toUpperCase();
                if (maThe) {
                    PropertiesService.getScriptProperties().setProperty("UNREAD_LIENHE_" + maThe, "true");
                }
                if (userSubmit && userSubmit !== maThe) {
                    PropertiesService.getScriptProperties().setProperty("UNREAD_LIENHE_" + userSubmit, "true");
                }
                return { success: true };
            }
        }
        return { success: false, message: "Không tìm thấy câu hỏi này" };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function markLienHeAsSeen(id) {
    try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LIEN_HE");
        if (!sheet) return { success: false };

        var data = sheet.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
            if (data[i][0] == id) {
                sheet.getRange(i + 1, 10).setValue(true);
                var maThe = String(data[i][2]).trim().toUpperCase();
                var userSubmit = String(data[i][10] || "").trim().toUpperCase();
                if (maThe) {
                    PropertiesService.getScriptProperties().setProperty("UNREAD_LIENHE_" + maThe, "false");
                }
                if (userSubmit && userSubmit !== maThe) {
                    PropertiesService.getScriptProperties().setProperty("UNREAD_LIENHE_" + userSubmit, "false");
                }
                return { success: true };
            }
        }
        return { success: false };
    } catch (e) {
        return { success: false };
    }
}

// ==========================================
// ĐA CÔNG ĐOẠN MODULE (BACKEND)
// ==========================================

function getDaCongDoanData() {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("DA_CONG_DOAN");
        if (!sheet) {
            sheet = ss.insertSheet("DA_CONG_DOAN");
            sheet.getRange("A1:B1").setValues([["KEY", "JSON_DATA"]]);
            return { success: true, data: {} };
        }

        var data = sheet.getDataRange().getValues();
        var result = {};
        for (var i = 1; i < data.length; i++) {
            var key = data[i][0];
            var val = data[i][1];
            if (key === "LINES_DATA" && val) {
                try {
                    result = JSON.parse(val);
                } catch (e) {
                    result = {};
                }
            }
        }
        return { success: true, data: result };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function saveDaCongDoanData(linesData) {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("DA_CONG_DOAN");
        if (!sheet) {
            sheet = ss.insertSheet("DA_CONG_DOAN");
            sheet.getRange("A1:B1").setValues([["KEY", "JSON_DATA"]]);
        }

        var data = sheet.getDataRange().getValues();
        var rowIndex = -1;
        for (var i = 1; i < data.length; i++) {
            if (data[i][0] === "LINES_DATA") {
                rowIndex = i + 1;
                break;
            }
        }

        var jsonStr = JSON.stringify(linesData);
        if (rowIndex === -1) {
            sheet.appendRow(["LINES_DATA", jsonStr]);
        } else {
            sheet.getRange(rowIndex, 2).setValue(jsonStr);
        }

        return { success: true };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function lookupDcdEmployees(idArray) {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var hrSheet = ss.getSheetByName("出勤数据");
        if (!hrSheet) {
            return { success: false, message: "Không tìm thấy sheet '出勤数据'." };
        }

        var hrData = hrSheet.getDataRange().getValues();
        var hrMap = {};
        for (var i = 1; i < hrData.length; i++) {
            var id = String(hrData[i][1]).trim().toUpperCase();
            var name = String(hrData[i][2]).trim();

            var statusStr = "";
            for (var col = 0; col < hrData[i].length; col++) {
                var cellVal = String(hrData[i][col]);
                if (cellVal.includes("Đã Nghỉ Việc") || cellVal.includes("已离职") || cellVal.includes("Nghỉ Việc") || cellVal.includes("Nghỉ việc")) {
                    statusStr = "Đã Nghỉ Việc";
                    break;
                }
            }
            if (!statusStr) statusStr = "Đang Làm Việc";

            if (id) {
                hrMap[id] = { name: name, status: statusStr };
            }
        }

        var result = [];
        for (var j = 0; j < idArray.length; j++) {
            var cleanId = String(idArray[j]).trim().toUpperCase();
            if (hrMap[cleanId]) {
                result.push({
                    id: cleanId,
                    name: hrMap[cleanId].name || "Chưa có tên",
                    status: hrMap[cleanId].status
                });
            } else {
                result.push({
                    id: cleanId,
                    name: "Không tìm thấy",
                    status: "Đang Làm Việc"
                });
            }
        }

        return { success: true, data: result };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function getDcdAdminData() {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("ADMIN_DATA");
        if (!sheet) return { success: false, message: "Không tìm thấy sheet ADMIN_DATA" };
        var lastRow = sheet.getLastRow();
        if (lastRow < 2) {
            return { success: true, data: { lines: [], leaders: [] } };
        }
        var data = sheet.getRange(2, 2, lastRow - 1, 2).getValues();
        var lines = [];
        var leaders = [];
        for (var i = 0; i < data.length; i++) {
            var l = String(data[i][0]).trim();
            var t = String(data[i][1]).trim();
            if (l && lines.indexOf(l) === -1) {
                lines.push(l);
            }
            if (t && leaders.indexOf(t) === -1) {
                leaders.push(t);
            }
        }
        return { success: true, data: { lines: lines, leaders: leaders } };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

// ==========================================
// PASS TRẠM MODULE (BACKEND)
// ==========================================

function getPassTramInitData() {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();

        // 1. Dữ liệu Dropdown từ sheet ADMIN_DATA
        var adminSheet = ss.getSheetByName("ADMIN_DATA");
        var adminData = adminSheet ? adminSheet.getDataRange().getValues() : [];
        var tramList = [], lineList = [], khuVucList = [], giaoVienList = [], modelList = [], ipqcList = [], meList = [];

        for (var i = 1; i < adminData.length; i++) {
            var tram = String(adminData[i][0] || "").trim();  // Cột A
            var line = String(adminData[i][1] || "").trim();  // Cột B
            var kv = String(adminData[i][3] || "").trim();    // Cột D
            var gv = String(adminData[i][18] || "").trim();   // Cột S (index 18)
            var model = String(adminData[i][5] || "").trim(); // Cột F
            var ipqc = String(adminData[i][16] || "").trim(); // Cột Q (index 16)
            var me = String(adminData[i][17] || "").trim();   // Cột R (index 17)

            if (tram && tramList.indexOf(tram) === -1) tramList.push(tram);
            if (line && lineList.indexOf(line) === -1) lineList.push(line);
            if (kv && khuVucList.indexOf(kv) === -1) khuVucList.push(kv);
            if (gv && giaoVienList.indexOf(gv) === -1) giaoVienList.push(gv);
            if (model && modelList.indexOf(model) === -1) modelList.push(model);
            if (ipqc && ipqcList.indexOf(ipqc) === -1) ipqcList.push(ipqc);
            if (me && meList.indexOf(me) === -1) meList.push(me);
        }

        // 2. Map Mã Thẻ -> Họ Và Tên từ sheet 出勤数据
        var hrSheet = ss.getSheetByName("出勤数据");
        var hrMap = {};
        if (hrSheet) {
            var hrData = hrSheet.getDataRange().getValues();
            for (var j = 1; j < hrData.length; j++) {
                var maThe = String(hrData[j][1] || "").trim().toUpperCase();
                var hoTen = String(hrData[j][2] || "").trim();
                if (maThe && !hrMap[maThe]) {
                    hrMap[maThe] = hoTen;
                }
            }
        }

        // 3. Danh sách Sheet khả dụng ở Google Sheet mục tiêu
        var targetSsId = "1LjPv9qkzUAsfWvoV-3Vifr-xSltbT1ufJPc1iwD_c4k";
        var targetSheets = ["STA5", "CFN4"];
        try {
            var targetSs = SpreadsheetApp.openById(targetSsId);
            var sheets = targetSs.getSheets();
            if (sheets && sheets.length) {
                targetSheets = sheets.map(function (s) { return s.getName(); });
            }
        } catch (errSheets) { }

        return {
            success: true,
            dropdowns: {
                tram: tramList,
                line: lineList,
                khuVuc: khuVucList,
                giaoVien: giaoVienList,
                model: modelList,
                ipqc: ipqcList,
                me: meList
            },
            hrMap: hrMap,
            targetSheets: targetSheets
        };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function savePassTramData(rowsData, targetSheetName) {
    var lock = LockService.getScriptLock();
    try {
        lock.tryLock(15000);
        if (!rowsData || !rowsData.length) {
            return { success: false, message: "Không có dữ liệu để lưu." };
        }

        targetSheetName = (targetSheetName || "STA5").trim();

        var targetSsId = "1LjPv9qkzUAsfWvoV-3Vifr-xSltbT1ufJPc1iwD_c4k";
        var targetSs = SpreadsheetApp.openById(targetSsId);
        var sheet = targetSs.getSheetByName(targetSheetName);
        if (!sheet) {
            sheet = targetSs.insertSheet(targetSheetName);
            sheet.getRange(1, 1, 1, 16).setValues([[
                "STT", "Model", "Khu Vực", "Line", "Ca",
                "Trạm CTQ", "Mã Thẻ", "Họ Và Tên", "Ngày Làm Việc",
                "Số Lượng (pcs)", "Kết quả", "Giáo Viên", "IPQC", "ME", "Note", "Số Giấy Tờ"
            ]]);
        }

        // Xác định dòng trống kế tiếp dựa trên cột B (Model)
        var colBValues = sheet.getRange("B:B").getValues();
        var lastRow = 0;
        for (var i = colBValues.length - 1; i >= 0; i--) {
            if (colBValues[i][0] !== "" && colBValues[i][0] !== null && colBValues[i][0] !== undefined) {
                lastRow = i + 1;
                break;
            }
        }
        var startRow = Math.max(2, lastRow + 1);

        // Tính Số Giấy Tờ mới (tăng tự động từ 1)
        var colPValues = sheet.getRange("P:P").getValues();
        var maxDocNo = 0;
        for (var j = 1; j < colPValues.length; j++) {
            var val = colPValues[j][0];
            if (val !== "" && val !== null && val !== undefined) {
                var num = parseInt(val, 10);
                if (!isNaN(num) && num > maxDocNo) {
                    maxDocNo = num;
                }
            }
        }
        var nextDocNo = maxDocNo + 1;

        // Chuẩn bị dữ liệu 14 cột từ cột B (Model) đến cột O (Note), bỏ cột STT (r[0])
        var formattedRows = rowsData.map(function (r) {
            return [
                r[1] || "", // Model (Cột B)
                r[2] || "", // Khu Vực (Cột C)
                r[3] || "", // Line (Cột D)
                r[4] || "", // Ca (Cột E)
                r[5] || "", // Trạm CTQ (Cột F)
                r[6] || "", // Mã Thẻ (Cột G)
                r[7] || "", // Họ Và Tên (Cột H)
                r[8] || "", // Ngày Làm Việc (Cột I)
                r[9] || "", // Số Lượng (pcs) (Cột J)
                r[10] || "", // Trạng Thái / Kết quả (PASS/FAIL) (Cột K)
                r[11] || "", // Giáo Viên (Cột L)
                r[12] || "", // IPQC (Cột M)
                r[13] || "", // ME (Cột N)
                r[14] || ""  // Note (Cột O)
            ];
        });

        // Ghi dữ liệu vào dải ô từ cột B (cột 2) đến cột O (cột 15, tổng 14 cột)
        var targetRange = sheet.getRange(startRow, 2, formattedRows.length, 14);
        targetRange.setValues(formattedRows);
        targetRange.setHorizontalAlignment("center").setVerticalAlignment("middle");

        // Ghi Số Giấy Tờ vào cột P (cột 16) & Gộp các ô vừa dán nếu lưu nhiều dòng
        var colPRange = sheet.getRange(startRow, 16, formattedRows.length, 1);
        if (formattedRows.length > 1) {
            colPRange.merge();
        }
        colPRange.setValue(nextDocNo)
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle")
            .setFontWeight("bold")
            .setFontSize(14);

        SpreadsheetApp.flush();

        return { success: true, count: formattedRows.length, startRow: startRow, docNo: nextDocNo };
    } catch (e) {
        return { success: false, message: e.toString() };
    } finally {
        try { lock.releaseLock(); } catch (err) { }
    }
}

function getPassTramReportData(sheetName, fromDate, toDate, maTheFilter, khuVucFilter, caFilter) {
    try {
        sheetName = (sheetName || "STA5").trim();
        var targetSsId = "1LjPv9qkzUAsfWvoV-3Vifr-xSltbT1ufJPc1iwD_c4k";
        var targetSs = SpreadsheetApp.openById(targetSsId);
        var sheet = targetSs.getSheetByName(sheetName);
        if (!sheet) {
            return { success: true, headers: [], rows: [] };
        }

        var values = sheet.getDataRange().getValues();
        if (!values || values.length <= 1) {
            return { success: true, headers: values[0] || [], rows: [] };
        }

        var headers = values[0];
        var allRows = values.slice(1);
        var filteredRows = [];

        var parseDateVal = function (dVal) {
            if (!dVal) return null;
            if (dVal instanceof Date) return dVal;
            var str = String(dVal).trim();
            if (str.indexOf("-") !== -1) {
                var p = str.split("-");
                if (p.length === 3) return new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
            } else if (str.indexOf("/") !== -1) {
                var p = str.split("/");
                if (p.length === 3) return new Date(parseInt(p[2], 10), parseInt(p[1], 10) - 1, parseInt(p[0], 10));
            }
            return null;
        };

        var fromD = parseDateVal(fromDate);
        var toD = parseDateVal(toDate);
        if (fromD) fromD.setHours(0, 0, 0, 0);
        if (toD) toD.setHours(23, 59, 59, 999);

        var searchMaThe = (maTheFilter || "").toString().trim().toUpperCase();
        var searchKhuVuc = (khuVucFilter || "").toString().trim().toUpperCase();
        var searchCa = (caFilter || "").toString().trim().toUpperCase();

        for (var i = 0; i < allRows.length; i++) {
            var row = allRows[i];
            var rawDate = row[8]; // Cột I (index 8)
            var rowMaThe = String(row[6] || "").trim().toUpperCase(); // Cột G (index 6)
            var rowKhuVuc = String(row[2] || "").trim().toUpperCase(); // Cột C (index 2)
            var rowCa = String(row[4] || "").trim().toUpperCase(); // Cột E (index 4)

            var rowDateObj = parseDateVal(rawDate);
            var dateStr = "";
            if (rawDate instanceof Date) {
                var y = rawDate.getFullYear();
                var m = ("0" + (rawDate.getMonth() + 1)).slice(-2);
                var d = ("0" + rawDate.getDate()).slice(-2);
                dateStr = d + "/" + m + "/" + y;
            } else {
                dateStr = String(rawDate || "").trim();
            }

            var matchDate = true;
            if (fromD || toD) {
                if (!rowDateObj) {
                    matchDate = false;
                } else {
                    if (fromD && rowDateObj < fromD) matchDate = false;
                    if (toD && rowDateObj > toD) matchDate = false;
                }
            }

            var matchMaThe = !searchMaThe || rowMaThe.indexOf(searchMaThe) !== -1;
            var matchKhuVuc = !searchKhuVuc || rowKhuVuc.indexOf(searchKhuVuc) !== -1;
            var matchCa = !searchCa || searchCa === "ALL" || rowCa.indexOf(searchCa) !== -1;

            if (matchDate && matchMaThe && matchKhuVuc && matchCa) {
                row[8] = dateStr;
                filteredRows.push(row);
            }
        }

        return { success: true, headers: headers, rows: filteredRows };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function getPassTramStatData(sheetName, fromDate, toDate, khuVuc, line, ca) {
    try {
        sheetName = (sheetName || "STA5").trim();
        var targetSsId = "1LjPv9qkzUAsfWvoV-3Vifr-xSltbT1ufJPc1iwD_c4k";
        var targetSs = SpreadsheetApp.openById(targetSsId);
        var sheet = targetSs.getSheetByName(sheetName);
        if (!sheet) {
            return { success: true, rows: [] };
        }

        var values = sheet.getDataRange().getValues();
        if (!values || values.length <= 1) {
            return { success: true, rows: [] };
        }

        var allRows = values.slice(1);
        var filteredRows = [];
        var khuVucMap = {};
        var lineMap = {};

        var fKv = (khuVuc || "ALL").trim().toUpperCase();
        var fLine = (line || "ALL").trim().toUpperCase();
        var fCa = (ca || "ALL").trim().toUpperCase();

        function parseToYMD(raw) {
            if (!raw) return "";
            if (raw instanceof Date) {
                var y = raw.getFullYear();
                var m = ("0" + (raw.getMonth() + 1)).slice(-2);
                var d = ("0" + raw.getDate()).slice(-2);
                return y + "-" + m + "-" + d;
            }
            var str = String(raw).trim();
            if (str.indexOf("/") !== -1) {
                var p = str.split("/");
                if (p.length === 3) {
                    var day = ("0" + p[0]).slice(-2);
                    var month = ("0" + p[1]).slice(-2);
                    var year = p[2];
                    if (year.length === 2) year = "20" + year;
                    return year + "-" + month + "-" + day;
                }
            }
            return str;
        }

        var ymdFrom = parseToYMD(fromDate);
        var ymdTo = parseToYMD(toDate);

        for (var i = 0; i < allRows.length; i++) {
            var r = allRows[i];
            var rModel = String(r[1] || "").trim();
            var rKhuVuc = String(r[2] || "").trim();
            var rLine = String(r[3] || "").trim();
            var rCa = String(r[4] || "").trim();
            var rTram = String(r[5] || "").trim();
            var rDateYMD = parseToYMD(r[8]);

            // Chỉ thu thập Khu Vực & Line từ các dòng có trạm CTQ thực tế (bỏ hàng trống)
            if (rTram && rTram !== 'Trạm Khác' && rTram !== 'Trạm Chưa Phân Loại') {
                if (rKhuVuc) khuVucMap[rKhuVuc] = true;
                if (rLine) lineMap[rLine] = true;
            }

            var matchKv = (fKv === "ALL" || rKhuVuc.toUpperCase() === fKv);
            var matchLine = (fLine === "ALL" || rLine.toUpperCase() === fLine);
            var matchCa = (fCa === "ALL" || rCa.toUpperCase() === fCa);

            var matchDate = true;
            if (ymdFrom && rDateYMD && rDateYMD < ymdFrom) matchDate = false;
            if (ymdTo && rDateYMD && rDateYMD > ymdTo) matchDate = false;

            if (matchKv && matchLine && matchCa && matchDate) {
                filteredRows.push({
                    model: rModel,
                    khuVuc: rKhuVuc,
                    line: rLine,
                    ca: rCa,
                    tram: rTram,
                    maThe: String(r[6] || "").trim(),
                    hoTen: String(r[7] || "").trim(),
                    ngay: rDateYMD,
                    pcs: Number(r[9] || 0),
                    ketQua: String(r[10] || "").trim().toUpperCase(),
                    giaoVien: String(r[11] || "").trim(),
                    ipqc: String(r[12] || "").trim(),
                    me: String(r[13] || "").trim()
                });
            }
        }

        return {
            success: true,
            rows: filteredRows,
            khuVucList: Object.keys(khuVucMap).sort(),
            lineList: Object.keys(lineMap).sort()
        };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function savePassTramSummaryData(matrixData, bgColorsMatrix, mergesList, hiddenColsList) {
    var lock = LockService.getScriptLock();
    try {
        lock.tryLock(15000);
        if (!matrixData || !matrixData.length) {
            return { success: false, message: "Không có dữ liệu bảng Summary để lưu." };
        }

        var sheetName = "Sumarry";
        var targetSsId = "1LjPv9qkzUAsfWvoV-3Vifr-xSltbT1ufJPc1iwD_c4k";
        var targetSs = SpreadsheetApp.openById(targetSsId);
        var sheet = targetSs.getSheetByName(sheetName);
        if (!sheet) {
            sheet = targetSs.insertSheet(sheetName);
        }

        sheet.clear();

        var numRows = matrixData.length;
        var numCols = matrixData[0].length;
        var range = sheet.getRange(1, 1, numRows, numCols);

        // 1. Set values
        try {
            range.setValues(matrixData);
        } catch (eV) { }

        // 2. Set font, alignment, wrap & border
        try {
            range.setFontFamily("Times New Roman");
            range.setFontSize(18);
            range.setFontWeight("bold");
            range.setHorizontalAlignment("center");
            range.setVerticalAlignment("middle");
            range.setWrap(true);
            range.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
        } catch (eF) { }

        // 3. Set background colors matrix
        if (bgColorsMatrix && bgColorsMatrix.length === numRows) {
            try {
                range.setBackgrounds(bgColorsMatrix);
            } catch (eBg) {
                for (var rIdx = 0; rIdx < numRows; rIdx++) {
                    if (bgColorsMatrix[rIdx] && bgColorsMatrix[rIdx].length === numCols) {
                        try {
                            sheet.getRange(rIdx + 1, 1, 1, numCols).setBackgrounds([bgColorsMatrix[rIdx]]);
                        } catch (eRowBg) { }
                    }
                }
            }
        }

        // 4. Merge cells
        if (mergesList && mergesList.length) {
            mergesList.forEach(function (m) {
                try {
                    sheet.getRange(m.r, m.c, m.rs, m.cs).merge();
                } catch (eM) { }
            });
        }

        // 5. Set Column Widths & Row Heights
        try {
            sheet.setColumnWidth(1, 120);
            sheet.setColumnWidth(2, 350);
            for (var c = 3; c <= numCols; c++) {
                sheet.setColumnWidth(c, 110);
            }

            sheet.setRowHeight(1, 65);
            for (var r = 2; r <= numRows; r++) {
                sheet.setRowHeight(r, 45);
            }
        } catch (eDim) { }

        // 6. Unhide all columns first
        try {
            sheet.showColumns(1, numCols);
        } catch (eShowAll) {
            try {
                range.unhideColumn();
            } catch (eUn) { }
        }

        // 7. Hide specific columns in hiddenColsList
        if (hiddenColsList && hiddenColsList.length) {
            hiddenColsList.forEach(function (colIdx) {
                if (colIdx >= 1 && colIdx <= numCols) {
                    try {
                        sheet.hideColumns(colIdx);
                    } catch (eH) { }
                }
            });
        }

        SpreadsheetApp.flush();

        return { success: true, rowsCount: numRows, colsCount: numCols, sheetName: sheetName };
    } catch (e) {
        return { success: false, message: e.toString() };
    } finally {
        try { lock.releaseLock(); } catch (err) { }
    }
}

function getPassTramSummarySheetDates() {
    try {
        var sheetName = "Sumarry";
        var targetSsId = "1LjPv9qkzUAsfWvoV-3Vifr-xSltbT1ufJPc1iwD_c4k";
        var targetSs = SpreadsheetApp.openById(targetSsId);
        var sheet = targetSs.getSheetByName(sheetName);
        if (!sheet) return { success: false };

        var lastCol = sheet.getLastColumn();
        if (lastCol < 3) return { success: false };

        var row2 = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
        var dates = [];
        var year = new Date().getFullYear();

        for (var c = 2; c < lastCol; c += 2) {
            var val = (row2[c] || '').toString().trim();
            if (val && val.indexOf('/') !== -1) {
                var parts = val.split('/');
                if (parts.length >= 2) {
                    var m = parseInt(parts[0], 10);
                    var d = parseInt(parts[1], 10);
                    if (!isNaN(m) && !isNaN(d)) {
                        dates.push({ year: year, month: m, day: d });
                    }
                }
            }
        }

        if (dates.length > 0) {
            return { success: true, dates: dates };
        }
        return { success: false };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function savePassTramTrainingData(matrixData, bgColorsMatrix, mergesList, hiddenColsList, customSheetName) {
    var lock = LockService.getScriptLock();
    try {
        lock.tryLock(15000);
        if (!matrixData || !matrixData.length) {
            return { success: false, message: "Không có dữ liệu bảng Training để lưu." };
        }

        var sheetName = customSheetName || "Training";
        var targetSsId = "1LjPv9qkzUAsfWvoV-3Vifr-xSltbT1ufJPc1iwD_c4k";
        var targetSs = SpreadsheetApp.openById(targetSsId);
        var sheet = targetSs.getSheetByName(sheetName);
        if (!sheet) {
            sheet = targetSs.insertSheet(sheetName);
        }

        sheet.clear();

        var numRows = matrixData.length;
        var numCols = matrixData[0].length;
        var range = sheet.getRange(1, 1, numRows, numCols);

        // 1. Set values
        try {
            range.setValues(matrixData);
        } catch (eV) { }

        // 2. Set font, alignment, wrap & border
        try {
            range.setFontFamily("Times New Roman");
            range.setFontSize(14);
            range.setFontWeight("bold");
            range.setHorizontalAlignment("center");
            range.setVerticalAlignment("middle");
            range.setWrap(true);
            range.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
        } catch (eF) { }

        // 3. Set background colors matrix
        if (bgColorsMatrix && bgColorsMatrix.length === numRows) {
            try {
                range.setBackgrounds(bgColorsMatrix);
            } catch (eBg) {
                for (var rIdx = 0; rIdx < numRows; rIdx++) {
                    if (bgColorsMatrix[rIdx] && bgColorsMatrix[rIdx].length === numCols) {
                        try {
                            sheet.getRange(rIdx + 1, 1, 1, numCols).setBackgrounds([bgColorsMatrix[rIdx]]);
                        } catch (eRowBg) { }
                    }
                }
            }
        }

        // 4. Merge cells
        if (mergesList && mergesList.length) {
            mergesList.forEach(function (m) {
                try {
                    sheet.getRange(m.r, m.c, m.rs, m.cs).merge();
                } catch (eM) { }
            });
        }

        // 5. Set Column Widths & Row Heights
        try {
            sheet.setColumnWidth(1, 100);
            sheet.setColumnWidth(2, 380);
            sheet.setColumnWidth(3, 80);
            sheet.setColumnWidth(4, 80);
            sheet.setColumnWidth(5, 80);
            for (var c = 6; c <= numCols; c++) {
                sheet.setColumnWidth(c, 70);
            }

            for (var r = 1; r <= numRows; r++) {
                sheet.setRowHeight(r, 38);
            }
        } catch (eDim) { }

        // 6. Unhide all columns first
        try {
            sheet.showColumns(1, numCols);
        } catch (eShowAll) {
            try {
                range.unhideColumn();
            } catch (eUn) { }
        }

        // 7. Hide specific columns in hiddenColsList
        if (hiddenColsList && hiddenColsList.length) {
            hiddenColsList.forEach(function (colIdx) {
                if (colIdx >= 1 && colIdx <= numCols) {
                    try {
                        sheet.hideColumns(colIdx);
                    } catch (eH) { }
                }
            });
        }

        SpreadsheetApp.flush();

        return { success: true, rowsCount: numRows, colsCount: numCols, sheetName: sheetName };
    } catch (e) {
        return { success: false, message: e.toString() };
    } finally {
        try { lock.releaseLock(); } catch (err) { }
    }
}

function getPassTramTrainingSheetDates() {
    try {
        var sheetName = "Training";
        var targetSsId = "1LjPv9qkzUAsfWvoV-3Vifr-xSltbT1ufJPc1iwD_c4k";
        var targetSs = SpreadsheetApp.openById(targetSsId);
        var sheet = targetSs.getSheetByName(sheetName);
        if (!sheet) return { success: false };

        var lastCol = sheet.getLastColumn();
        if (lastCol < 6) return { success: false };

        var row2 = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
        var dates = [];
        var year = new Date().getFullYear();

        for (var c = 5; c < lastCol; c += 2) {
            var val = (row2[c] || '').toString().trim();
            if (val && val.indexOf('/') !== -1) {
                var parts = val.split('/');
                if (parts.length >= 2) {
                    var d = parseInt(parts[0], 10);
                    var m = parseInt(parts[1], 10);
                    if (!isNaN(m) && !isNaN(d)) {
                        dates.push({ year: year, month: m, day: d });
                    }
                }
            }
        }

        if (dates.length > 0) {
            return { success: true, dates: dates };
        }
        return { success: false };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function getPassTramKhuVucList(sheetName) {
    try {
        sheetName = (sheetName || "STA5").trim();
        var targetSsId = "1LjPv9qkzUAsfWvoV-3Vifr-xSltbT1ufJPc1iwD_c4k";
        var targetSs = SpreadsheetApp.openById(targetSsId);
        var sheet = targetSs.getSheetByName(sheetName);
        if (!sheet) return { success: true, list: [] };

        var values = sheet.getDataRange().getValues();
        if (!values || values.length <= 1) return { success: true, list: [] };

        var khuVucSet = {};
        var list = [];
        for (var i = 1; i < values.length; i++) {
            var kv = String(values[i][2] || "").trim();
            if (kv && !khuVucSet[kv]) {
                khuVucSet[kv] = true;
                list.push(kv);
            }
        }
        return { success: true, list: list };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function lookupEmployeeNameDcd(maThe) {
    try {
        if (!maThe) return { success: false, name: "" };
        var searchCode = String(maThe).trim().toUpperCase();
        var ss = SpreadsheetApp.getActiveSpreadsheet();

        // 1. Kiểm tra sheet 出勤数据 trước tiên (Cột C: Họ Tên)
        var hrSheet = ss.getSheetByName("出勤数据");
        if (hrSheet) {
            var dataH = hrSheet.getDataRange().getDisplayValues();
            for (var i = 1; i < dataH.length; i++) {
                var r = dataH[i];
                for (var c = 0; c < r.length; c++) {
                    if (String(r[c] || "").trim().toUpperCase() === searchCode) {
                        var nameC = String(r[2] || "").trim(); // Cột C
                        if (nameC && nameC.toUpperCase() !== searchCode) return { success: true, name: nameC };
                        var alt = String(r[1] || r[3] || r[0] || "").trim();
                        if (alt && alt.toUpperCase() !== searchCode) return { success: true, name: alt };
                    }
                }
            }
        }

        // 2. Kiểm tra sheet CHUYÊN CẦN STA5
        var sta5Sheet = ss.getSheetByName("CHUYÊN CẦN STA5");
        if (sta5Sheet) {
            var dataSta = sta5Sheet.getDataRange().getDisplayValues();
            for (var i = 1; i < dataSta.length; i++) {
                if (String(dataSta[i][0] || "").trim().toUpperCase() === searchCode) {
                    var nameSta = String(dataSta[i][1] || "").trim();
                    if (nameSta) return { success: true, name: nameSta };
                }
            }
        }

        // 3. Kiểm tra sheet NHẬP THẺ / CẤP THẺ
        var nhapTheSheet = ss.getSheetByName("NHẬP THẺ") || ss.getSheetByName("CẤP THẺ");
        if (nhapTheSheet) {
            var dataN = nhapTheSheet.getDataRange().getDisplayValues();
            for (var i = 1; i < dataN.length; i++) {
                var codeN = String(dataN[i][2] || dataN[i][1] || "").trim().toUpperCase();
                if (codeN === searchCode) {
                    var nameN = String(dataN[i][3] || dataN[i][2] || "").trim();
                    if (nameN) return { success: true, name: nameN };
                }
            }
        }

        return { success: false, name: "" };
    } catch (e) {
        return { success: false, name: "", message: e.toString() };
    }
}


function getAdminDataForDcd() {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("ADMIN_DATA");
        if (!sheet) return { success: false, message: "Không tìm thấy sheet ADMIN_DATA" };

        var data = sheet.getDataRange().getValues();
        var models = [], lines = [], leaders = [], stations = [], teachers = [];

        for (var i = 1; i < data.length; i++) {
            var station = String(data[i][0] || "").trim();
            var line = String(data[i][1] || "").trim();
            var leader = String(data[i][2] || "").trim();
            var model = String(data[i][5] || "").trim();
            var teacher = String(data[i][18] || "").trim();

            if (station && stations.indexOf(station) === -1) stations.push(station);
            if (line && lines.indexOf(line) === -1) lines.push(line);
            if (leader && leaders.indexOf(leader) === -1) leaders.push(leader);
            if (model && models.indexOf(model) === -1) models.push(model);
            if (teacher && teachers.indexOf(teacher) === -1) teachers.push(teacher);
        }

        return {
            success: true,
            models: models,
            lines: lines,
            leaders: leaders,
            stations: stations,
            teachers: teachers
        };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function saveDcdRecord(records) {
    try {
        if (!records) return { success: false, message: "Dữ liệu không hợp lệ" };
        if (!Array.isArray(records[0])) {
            records = [records];
        }

        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = getDcdSheet_();
        if (!sheet) {
            sheet = ss.insertSheet("DS ĐA CÔNG ĐOẠN");
            sheet.getRange("A1:L1").setValues([["MODEL", "Mã Thẻ", "Họ Và Tên", "Line", "Tổ trưởng", "Vị Trí Ban Đầu", "Giáo Viên", "PASS/FAIL", "Ngày", "Trạm CTQ 1", "Trạm CTQ 2", "Trạm CTQ 3"]]);
            sheet.getRange("A1:L1").setFontWeight("bold").setBackground("#f3f4f6");
        }

        var maxRows = sheet.getMaxRows();
        var dataB = sheet.getRange("B1:B").getValues();
        var emptyRow = dataB.length + 1;

        // Dò từ trên xuống (từ dòng 2) xem dòng nào trống (Cột B - Mã Thẻ rỗng)
        for (var i = 1; i < dataB.length; i++) {
            if (!String(dataB[i][0]).trim()) {
                emptyRow = i + 1;
                break;
            }
        }

        // Nếu ghi vượt quá số dòng tối đa của sheet thì thêm dòng mới
        if (emptyRow + records.length - 1 > maxRows) {
            sheet.insertRowsAfter(maxRows, records.length);
        }

        // Ghi dữ liệu vào vị trí trống đầu tiên
        sheet.getRange(emptyRow, 1, records.length, records[0].length).setValues(records);

        return { success: true };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('🚀 BÁO CÁO CHUYÊN CẦN')
        .addItem('Đồng bộ dữ liệu Ngày/Ca', 'showDialog')
        .addToUi();
}

function showDialog() {
    var html = HtmlService.createHtmlOutputFromFile('CTQ2026_copythuong')
        .setWidth(360)
        .setHeight(300)
        .setTitle('Đồng Bộ Dữ Liệu Chuyên Cần');
    SpreadsheetApp.getUi().showModalDialog(html, 'Đồng bộ từ File Tổng');
}

function normalizeCa(caStr) {
    if (!caStr) return "";
    var s = String(caStr).trim().toUpperCase();
    if (s === "D" || s === "DAY" || s === "CA NGÀY" || s.indexOf("D") === 0) return "DAY";
    if (s === "N" || s === "NIGHT" || s === "CA ĐÊM" || s.indexOf("N") === 0) return "NIGHT";
    return s;
}

function normalizeKey(str) {
    if (!str) return "";
    return String(str).trim().toUpperCase();
}

function extractSpreadsheetId(urlOrId) {
    if (!urlOrId) return "";
    urlOrId = String(urlOrId).trim();
    var match = urlOrId.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (match && match[1]) return match[1];
    return urlOrId;
}

function processCopyData(datesInput, selectedCa, targetUrl, customSheetName, forcePaste, sourceUrl, customSourceSheetName) {
    try {
        var sourceFileId = extractSpreadsheetId(sourceUrl) || "1_T-OtIZF7_dVXZzTxtXzdLTr5KtWL642SD7kGk8fEH0"; // File Nguồn
        var sourceSs = SpreadsheetApp.openById(sourceFileId);
        var reqSrcSheetName = (customSourceSheetName && String(customSourceSheetName).trim() !== "") ? String(customSourceSheetName).trim() : "CHUYÊN CẦN STA5";
        var sourceSheet = sourceSs.getSheetByName(reqSrcSheetName) || sourceSs.getSheetByName("CHUYÊN CẦN STA5") || sourceSs.getActiveSheet();

        if (!sourceSheet) {
            return { success: false, message: "Không tìm thấy Sheet '" + reqSrcSheetName + "' ở File Nguồn!" };
        }

        var srcSheetNameActual = sourceSheet.getName();

        var targetFileId = extractSpreadsheetId(targetUrl) || "1zURB1fThVmD1bNY0o6XiDUMsTjLu9IjE4H7sZqF1Ing"; // File Đích
        var targetSs;
        try {
            targetSs = SpreadsheetApp.openById(targetFileId);
        } catch (e) {
            targetSs = SpreadsheetApp.getActiveSpreadsheet();
        }

        if (!targetSs) {
            return { success: false, message: "Không mở được File Đích! Vui lòng kiểm tra lại Link hoặc quyền truy cập." };
        }

        // Convert datesInput to array if string
        var datesArray = [];
        if (Array.isArray(datesInput)) {
            datesArray = datesInput;
        } else if (typeof datesInput === 'string') {
            datesArray = datesInput.split(',').map(function (s) { return s.trim(); }).filter(Boolean);
        }

        if (datesArray.length === 0) {
            return { success: false, message: "Vui lòng chọn ít nhất 1 ngày!" };
        }

        var srcHeaderDate = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
        var lastSrcRow = sourceSheet.getLastRow();

        var srcIds = sourceSheet.getRange(2, 1, lastSrcRow - 1, 1).getValues();      // Cột A: Mã thẻ
        var srcNames = sourceSheet.getRange(2, 2, lastSrcRow - 1, 1).getValues();    // Cột B: Họ Tên
        var srcStations = sourceSheet.getRange(2, 3, lastSrcRow - 1, 1).getValues(); // Cột C: Trạm
        var srcCas = sourceSheet.getRange(2, 4, lastSrcRow - 1, 1).getValues();      // Cột D: Ca (D/N)
        var srcLines = sourceSheet.getRange(2, 5, lastSrcRow - 1, 1).getValues();    // Cột E: Line

        // Tìm Sheet Đích để dùng cho pre-check & dán
        var datePartsInit = datesArray[0].split('/');
        var monthNumInit = parseInt(datePartsInit[1], 10) || 8;
        var targetSheetCheck = null;
        if (customSheetName && String(customSheetName).trim() !== "") {
            targetSheetCheck = targetSs.getSheetByName(String(customSheetName).trim());
        }
        if (!targetSheetCheck) {
            targetSheetCheck = targetSs.getSheetByName("CHUYỂN CẦN THÁNG " + monthNumInit) ||
                targetSs.getSheetByName("CHUYÊN CẦN THÁNG " + monthNumInit) ||
                targetSs.getActiveSheet();
        }
        var tgtSheetNameActual = targetSheetCheck ? targetSheetCheck.getName() : "Sheet Đích";

        // -------------------------------------------------------------
        // PRE-CHECK THÔNG MINH: CẢNH BÁO LỆCH SỐ LƯỢNG HÀNG THEO TRẠM
        // -------------------------------------------------------------
        if (!forcePaste) {
            var srcStationCounts = {};
            for (var sIdx = 0; sIdx < srcIds.length; sIdx++) {
                var sId = String(srcIds[sIdx][0]).trim();
                var sSt = String(srcStations[sIdx][0]).trim();
                var sCa = normalizeCa(srcCas[sIdx][0]);
                if (sId && sId !== "#N/A" && (selectedCa === "ALL" || sCa === normalizeCa(selectedCa))) {
                    var sKey = sSt + "|" + sCa;
                    srcStationCounts[sKey] = (srcStationCounts[sKey] || 0) + 1;
                }
            }

            var tgtStationCounts = {};
            if (targetSheetCheck) {
                var checkRows = targetSheetCheck.getLastRow() - 1;
                if (checkRows > 0) {
                    var tStations = targetSheetCheck.getRange(2, 6, checkRows, 1).getValues();
                    var tCas = targetSheetCheck.getRange(2, 7, checkRows, 1).getValues();
                    for (var tIdx = 0; tIdx < tStations.length; tIdx++) {
                        var tSt = String(tStations[tIdx][0]).trim();
                        var tCa = normalizeCa(tCas[tIdx][0]);
                        if (tSt && (selectedCa === "ALL" || tCa === normalizeCa(selectedCa))) {
                            var tKey = tSt + "|" + tCa;
                            tgtStationCounts[tKey] = (tgtStationCounts[tKey] || 0) + 1;
                        }
                    }
                }
            }

            var capacityWarnings = [];
            for (var stKey in srcStationCounts) {
                var srcCount = srcStationCounts[stKey];
                var tgtCount = tgtStationCounts[stKey] || 0;
                if (tgtCount > 0 && srcCount > tgtCount) {
                    var kParts = stKey.split('|');
                    capacityWarnings.push({
                        station: kParts[0],
                        ca: kParts[1] || selectedCa,
                        srcCount: srcCount,
                        tgtCount: tgtCount,
                        diff: srcCount - tgtCount
                    });
                }
            }

            if (capacityWarnings.length > 0) {
                return {
                    success: true,
                    capacityWarning: true,
                    capacityWarnings: capacityWarnings,
                    srcSheetName: srcSheetNameActual,
                    tgtSheetName: tgtSheetNameActual,
                    message: "⚠️ <b>CẢNH BÁO LỆCH SỐ LƯỢNG HÀNG THEO TRẠM:</b> Phát hiện Sheet " + srcSheetNameActual + " có số lượng người ở trạm nhiều hơn số hàng có sẵn trên Sheet " + tgtSheetNameActual + "!"
                };
            }
        }

        var processedDays = 0;
        var totalUpdatedRows = 0;
        var dayLogs = [];
        var mismatchMapByRow = {};

        for (var d = 0; d < datesArray.length; d++) {
            var targetDateStr = datesArray[d];

            // 1. Tìm cột ngày bên File Nguồn
            var srcDateColIndex = -1;
            for (var j = 0; j < srcHeaderDate.length; j++) {
                var val = srcHeaderDate[j];
                var formattedVal = "";
                if (val instanceof Date) {
                    formattedVal = Utilities.formatDate(val, sourceSs.getSpreadsheetTimeZone(), "dd/MM/yyyy");
                } else {
                    formattedVal = String(val).trim();
                }
                if (formattedVal === targetDateStr) {
                    srcDateColIndex = j + 1;
                    break;
                }
            }

            if (srcDateColIndex === -1) {
                dayLogs.push("⚠️ Ngày <b>" + targetDateStr + "</b>: Không tìm thấy trên Sheet " + srcSheetNameActual + "!");
                continue;
            }

            var srcValues = sourceSheet.getRange(2, srcDateColIndex, lastSrcRow - 1, 1).getValues();
            var srcBgColors = sourceSheet.getRange(2, srcDateColIndex, lastSrcRow - 1, 1).getBackgrounds();
            var srcFontColors = sourceSheet.getRange(2, srcDateColIndex, lastSrcRow - 1, 1).getFontColors();
            var srcNotes = sourceSheet.getRange(2, srcDateColIndex, lastSrcRow - 1, 1).getNotes();

            var mapById = {};
            var groupMap = {};

            for (var i = 0; i < srcIds.length; i++) {
                var id = String(srcIds[i][0]).trim();
                var name = String(srcNames[i][0]).trim();
                var station = String(srcStations[i][0]).trim();
                var caNorm = normalizeCa(srcCas[i][0]);

                var dataObj = {
                    id: id,
                    name: name,
                    station: station,
                    line: srcLines[i][0],
                    ca: srcCas[i][0],
                    value: srcValues[i][0],
                    bg: srcBgColors[i][0],
                    fontColor: srcFontColors[i][0],
                    note: srcNotes[i][0]
                };

                if (id && id !== "#N/A") {
                    mapById[id] = dataObj;
                }

                var groupKey = normalizeKey(station) + "|" + caNorm;
                if (!groupMap[groupKey]) {
                    groupMap[groupKey] = [];
                }
                groupMap[groupKey].push(dataObj);
            }

            // 2. Tìm Sheet Đích
            var dateParts = targetDateStr.split('/');
            var dayNum = parseInt(dateParts[0], 10);
            var monthNum = parseInt(dateParts[1], 10);

            var targetSheet = null;
            if (customSheetName && String(customSheetName).trim() !== "") {
                targetSheet = targetSs.getSheetByName(String(customSheetName).trim());
            }
            if (!targetSheet) {
                targetSheet = targetSs.getSheetByName("CHUYỂN CẦN THÁNG " + monthNum) ||
                    targetSs.getSheetByName("CHUYÊN CẦN THÁNG " + monthNum) ||
                    targetSs.getActiveSheet();
            }

            function findDateColInSheet(sh) {
                try {
                    var lastCol = sh.getLastColumn();
                    if (lastCol < 1) return -1;
                    var header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
                    for (var k = 0; k < header.length; k++) {
                        var cellVal = header[k];
                        if (cellVal instanceof Date) {
                            if (cellVal.getDate() === dayNum && (cellVal.getMonth() + 1) === monthNum) {
                                return k + 1;
                            }
                        } else if (cellVal !== "" && cellVal !== null && cellVal !== undefined) {
                            var strVal = String(cellVal).trim();
                            if (strVal === dayNum + "/" + monthNum || strVal === dateParts[0] + "/" + dateParts[1] || strVal === targetDateStr) {
                                return k + 1;
                            }
                        }
                    }
                } catch (err) { }
                return -1;
            }

            var tgtDateColIndex = findDateColInSheet(targetSheet);

            if (tgtDateColIndex === -1) {
                var allSheets = targetSs.getSheets();
                for (var s = 0; s < allSheets.length; s++) {
                    var colIdx = findDateColInSheet(allSheets[s]);
                    if (colIdx !== -1) {
                        targetSheet = allSheets[s];
                        tgtDateColIndex = colIdx;
                        break;
                    }
                }
            }

            if (tgtDateColIndex === -1) {
                dayLogs.push("⚠️ Ngày <b>" + targetDateStr + "</b>: Không tìm thấy cột ngày ở Sheet " + tgtSheetNameActual + "!");
                continue;
            }

            var lastTgtRow = targetSheet.getLastRow();
            var numRows = lastTgtRow - 1;
            if (numRows < 1) continue;

            var tgtIds = targetSheet.getRange(2, 4, numRows, 1).getValues();      // Col D
            var tgtNames = targetSheet.getRange(2, 5, numRows, 1).getValues();    // Col E
            var tgtStations = targetSheet.getRange(2, 6, numRows, 1).getValues(); // Col F
            var tgtCas = targetSheet.getRange(2, 7, numRows, 1).getValues();      // Col G
            var tgtLines = targetSheet.getRange(2, 8, numRows, 1).getValues();    // Col H

            var tgtValues = targetSheet.getRange(2, tgtDateColIndex, numRows, 1).getValues();
            var tgtBgs = targetSheet.getRange(2, tgtDateColIndex, numRows, 1).getBackgrounds();
            var tgtFonts = targetSheet.getRange(2, tgtDateColIndex, numRows, 1).getFontColors();
            var tgtNotes = targetSheet.getRange(2, tgtDateColIndex, numRows, 1).getNotes();

            var existingTargetIds = {};
            for (var idx = 0; idx < tgtIds.length; idx++) {
                var tid = String(tgtIds[idx][0]).trim();
                if (tid && tid !== "#N/A") existingTargetIds[tid] = true;
            }

            var outValues = [];
            var outBgs = [];
            var outFontColors = [];
            var outNotes = [];
            var assignedIdsInRun = {};
            var updatedCount = 0;

            for (var r = 0; r < tgtIds.length; r++) {
                var tgtId = String(tgtIds[r][0]).trim();
                var tgtName = String(tgtNames[r][0]).trim();
                var tgtStation = String(tgtStations[r][0]).trim();
                var tgtCaNorm = normalizeCa(tgtCas[r][0]);
                var tgtLineVal = String(tgtLines[r][0]).trim();

                var isCaMatched = (selectedCa === "ALL" || tgtCaNorm === normalizeCa(selectedCa));

                if (isCaMatched) {
                    var groupKey = normalizeKey(tgtStation) + "|" + tgtCaNorm;

                    // TRƯỜNG HỢP 1: Hàng ở Đích CÓ Mã Thẻ (tgtId hợp lệ)
                    if (tgtId && tgtId !== "#N/A") {
                        if (mapById.hasOwnProperty(tgtId)) {
                            var foundData = mapById[tgtId];
                            var srcLineNorm = normalizeKey(foundData.line);
                            var tgtLineNorm = normalizeKey(tgtLineVal);
                            var srcStationNorm = normalizeKey(foundData.station);
                            var tgtStationNorm = normalizeKey(tgtStation);

                            // KIỂM TRA NGHIÊM NGẶT ĐÚNG TRẠM VÀ ĐÚNG LINE (YÊU CẦU 1 & 3)
                            var isLineMatch = (srcLineNorm === "" || tgtLineNorm === "" || srcLineNorm === tgtLineNorm);
                            var isStationMatch = (srcStationNorm === "" || tgtStationNorm === "" || srcStationNorm === tgtStationNorm);

                            if (isLineMatch && isStationMatch) {
                                var valFromSrc = foundData.value;
                                if (valFromSrc === "" || valFromSrc === null || valFromSrc === undefined) {
                                    outValues.push(["NO"]);
                                    outBgs.push(["#ffffff"]);
                                    outFontColors.push(["#000000"]);
                                    outNotes.push([""]);
                                } else {
                                    outValues.push([valFromSrc]);
                                    outBgs.push([foundData.bg]);
                                    outFontColors.push([foundData.fontColor]);
                                    outNotes.push([foundData.note]);
                                    updatedCount++;
                                }
                            } else {
                                // KHÁC LINE HẶC KHÁC TRẠM -> ĐIỀN "Sai" (Nền đỏ) & BÁO LÊN BẢNG ĐỐI CHIẾU
                                outValues.push(["Sai"]);
                                outBgs.push(["#fee2e2"]);
                                outFontColors.push(["#dc2626"]);
                                var mismatchReason = !isLineMatch ? ("Khác Line (" + srcSheetNameActual + ": " + foundData.line + " ➔ " + tgtSheetNameActual + ": " + tgtLineVal + ")") : ("Khác Trạm (" + srcSheetNameActual + ": " + foundData.station + " ➔ " + tgtSheetNameActual + ": " + tgtStation + ")");
                                outNotes.push([""]);
                                var rowNum = r + 2;
                                dayLogs.push("⚠️ " + targetDateStr + " - Hàng <b>" + rowNum + "</b>: Mã thẻ <b>" + tgtId + " (" + tgtName + ")</b> " + mismatchReason + " ➔ Điền <b>Sai</b>");

                                // Tìm nhân sự ở Sheet Nguồn chuẩn đúng theo Trạm này (tgtStation) và Line này (tgtLineVal) để kéo sang khi bấm "Khớp"
                                var candForStation = null;
                                if (groupMap[groupKey]) {
                                    var members = groupMap[groupKey];
                                    for (var m = 0; m < members.length; m++) {
                                        var member = members[m];
                                        if (member.id && member.id !== "#N/A" && !assignedIdsInRun[member.id]) {
                                            var mLineNorm = normalizeKey(member.line);
                                            var tLineNorm = normalizeKey(tgtLineVal);
                                            if (mLineNorm === "" || tLineNorm === "" || mLineNorm === tLineNorm) {
                                                candForStation = member;
                                                assignedIdsInRun[member.id] = true;
                                                break;
                                            }
                                        }
                                    }
                                    if (!candForStation && members.length > 0) {
                                        for (var m2 = 0; m2 < members.length; m2++) {
                                            var member2 = members[m2];
                                            if (member2.id && member2.id !== "#N/A" && !assignedIdsInRun[member2.id]) {
                                                candForStation = member2;
                                                assignedIdsInRun[member2.id] = true;
                                                break;
                                            }
                                        }
                                    }
                                }

                                if (!mismatchMapByRow[rowNum]) {
                                    mismatchMapByRow[rowNum] = {
                                        targetRow: rowNum,
                                        srcId: candForStation ? candForStation.id : tgtId,
                                        srcName: candForStation ? candForStation.name : tgtName,
                                        srcStation: (candForStation ? candForStation.station : foundData.station) + " [Line " + (candForStation ? candForStation.line : foundData.line) + "]",
                                        tgtStation: tgtStation + " [Line " + tgtLineVal + "]",
                                        type: "MISMATCH_LINE_STATION",
                                        dates: []
                                    };
                                }
                                mismatchMapByRow[rowNum].dates.push({
                                    dateStr: targetDateStr,
                                    srcVal: candForStation ? (candForStation.value || "V") : "V",
                                    srcBg: candForStation ? (candForStation.bg || "#dcfce7") : "#dcfce7",
                                    srcFontColor: candForStation ? (candForStation.fontColor || "#000000") : "#000000",
                                    srcNote: candForStation ? (candForStation.note || "") : ""
                                });
                            }
                        } else {
                            // YÊU CẦU: Đích CÓ tên nhưng Nguồn KHÔNG CÓ -> Chỉ điền "Sai" (Nền đỏ) ở cột ngày, KHÔNG sửa Mã thẻ & Họ tên ở File Đích!
                            outValues.push(["Sai"]);
                            outBgs.push(["#fee2e2"]);
                            outFontColors.push(["#dc2626"]);
                            outNotes.push([""]);
                            var rowNum = r + 2;
                            dayLogs.push("⚠️ " + targetDateStr + " - Hàng <b>" + rowNum + "</b>: " + tgtSheetNameActual + " có <b>" + tgtId + " " + tgtName + "</b> nhưng " + srcSheetNameActual + " không có ➔ Điền <b>Sai</b>");

                            // Tìm nhân sự ở Sheet Nguồn thuộc cùng Trạm & Ca chưa được dán để sẵn sàng kéo sang khi bấm "Khớp"
                            var cand = null;
                            if (groupMap[groupKey]) {
                                var members = groupMap[groupKey];
                                for (var m = 0; m < members.length; m++) {
                                    var member = members[m];
                                    if (member.id && member.id !== "#N/A" && !assignedIdsInRun[member.id]) {
                                        var mLineNorm2 = normalizeKey(member.line);
                                        var tLineNorm2 = normalizeKey(tgtLineVal);
                                        if (mLineNorm2 === "" || tLineNorm2 === "" || mLineNorm2 === tLineNorm2) {
                                            cand = member;
                                            assignedIdsInRun[member.id] = true;
                                            break;
                                        }
                                    }
                                }
                                if (!cand && members.length > 0) {
                                    for (var m3 = 0; m3 < members.length; m3++) {
                                        var member3 = members[m3];
                                        if (member3.id && member3.id !== "#N/A" && !assignedIdsInRun[member3.id]) {
                                            cand = member3;
                                            assignedIdsInRun[member3.id] = true;
                                            break;
                                        }
                                    }
                                }
                            }

                            if (!mismatchMapByRow[rowNum]) {
                                mismatchMapByRow[rowNum] = {
                                    targetRow: rowNum,
                                    srcId: cand ? cand.id : ("Thiếu ở " + srcSheetNameActual),
                                    srcName: cand ? cand.name : ("(Chỉ có ở " + tgtSheetNameActual + ": " + tgtId + " " + tgtName + ")"),
                                    displaySrcId: cand ? cand.id : ("Thiếu ở " + srcSheetNameActual),
                                    displaySrcName: cand ? cand.name : ("(Chỉ có ở " + tgtSheetNameActual + ": " + tgtId + " " + tgtName + ")"),
                                    originalTgtId: tgtId,
                                    originalTgtName: tgtName,
                                    srcStation: cand ? cand.station : tgtStation,
                                    tgtStation: tgtStation + " [Line " + tgtLineVal + "]",
                                    type: "REVERSE_MISMATCH",
                                    dates: []
                                };
                            }
                            mismatchMapByRow[rowNum].dates.push({
                                dateStr: targetDateStr,
                                srcVal: cand ? (cand.value || "V") : "V",
                                srcBg: cand ? (cand.bg || "#dcfce7") : "#dcfce7",
                                srcFontColor: cand ? (cand.fontColor || "#000000") : "#000000",
                                srcNote: cand ? (cand.note || "") : ""
                            });
                        }
                    }
                    // TRƯỜNG HỢP 2: Hàng ở Đích THIẾU Mã Thẻ (Trống hoặc #N/A)
                    else {
                        var hasSrcPersonWithData = false;
                        var srcPersonInfo = "";
                        var matchedCandidate = null;

                        if (groupMap[groupKey]) {
                            var members = groupMap[groupKey];
                            for (var m = 0; m < members.length; m++) {
                                var candidate = members[m];
                                if (candidate.id && candidate.id !== "#N/A") {
                                    if (!existingTargetIds[candidate.id] && !assignedIdsInRun[candidate.id]) {
                                        if (candidate.value !== "" && candidate.value !== null && candidate.value !== undefined && candidate.value !== "NO") {
                                            hasSrcPersonWithData = true;
                                            matchedCandidate = candidate;
                                            srcPersonInfo = candidate.id + " " + candidate.name;
                                            existingTargetIds[candidate.id] = true;
                                            assignedIdsInRun[candidate.id] = true;
                                            break;
                                        }
                                    }
                                }
                            }
                        }

                        if (hasSrcPersonWithData && matchedCandidate) {
                            // Nguồn CÓ dữ liệu mà Đích KHÔNG CÓ Mã Thẻ -> Điền "Sai" (Nền đỏ)
                            outValues.push(["Sai"]);
                            outBgs.push(["#fee2e2"]);        // Nền đỏ nhạt
                            outFontColors.push(["#dc2626"]);   // Chữ đỏ đậm
                            outNotes.push([""]);
                            var rowNum = r + 2;
                            dayLogs.push("⚠️ " + targetDateStr + " - Hàng <b>" + rowNum + "</b>: " + srcSheetNameActual + " có <b>" + srcPersonInfo + "</b> nhưng " + tgtSheetNameActual + " thiếu Mã Thẻ ➔ Điền <b>Sai</b>");

                            if (!mismatchMapByRow[rowNum]) {
                                mismatchMapByRow[rowNum] = {
                                    targetRow: rowNum,
                                    srcId: matchedCandidate.id,
                                    srcName: matchedCandidate.name,
                                    srcStation: matchedCandidate.station,
                                    tgtStation: tgtStation,
                                    type: "NORMAL_MISMATCH",
                                    dates: []
                                };
                            }
                            mismatchMapByRow[rowNum].dates.push({
                                dateStr: targetDateStr,
                                srcVal: matchedCandidate.value,
                                srcBg: matchedCandidate.bg,
                                srcFontColor: matchedCandidate.fontColor,
                                srcNote: matchedCandidate.note || ""
                            });
                        } else {
                            outValues.push(["NO"]);
                            outBgs.push(["#ffffff"]);
                            outFontColors.push(["#000000"]);
                            outNotes.push([""]);
                        }
                    }
                } else {
                    outValues.push([tgtValues[r][0]]);
                    outBgs.push([tgtBgs[r][0]]);
                    outFontColors.push([tgtFonts[r][0]]);
                    outNotes.push([tgtNotes[r][0]]);
                }
            }

            var targetRange = targetSheet.getRange(2, tgtDateColIndex, numRows, 1);
            targetRange.setValues(outValues);
            targetRange.setBackgrounds(outBgs);
            targetRange.setFontColors(outFontColors);
            targetRange.setNotes(outNotes);

            processedDays++;
            totalUpdatedRows += updatedCount;
            dayLogs.push("✅ Ngày <b>" + targetDateStr + "</b>: Đồng bộ thành công <b>" + updatedCount + " dòng</b>.");
        }

        var mismatchList = [];
        for (var k in mismatchMapByRow) {
            if (mismatchMapByRow.hasOwnProperty(k)) {
                mismatchList.push(mismatchMapByRow[k]);
            }
        }

        var detailsHtml = "<div style='margin-top:12px; max-height:220px; overflow-y:auto; background:#ffffff; padding:12px; border-radius:8px; border:1px solid #cbd5e1; font-size:14px; color:#1e293b; text-align:left; box-shadow:inset 0 1px 3px rgba(0,0,0,0.05);'>" +
            "<b style='color:#1d4ed8;'>📌 CHI TIẾT KẾT QUẢ ĐỒNG BỘ NHIỀU NGÀY:</b><br><br>" +
            dayLogs.join("<br>") +
            "</div>";

        return {
            success: true,
            message: "Hoàn tất đồng bộ <b>" + processedDays + " ngày</b> (Tổng cộng <b>" + totalUpdatedRows + " lượt dòng</b>)." + detailsHtml,
            mismatchList: mismatchList
        };

    } catch (error) {
        return { success: false, message: "Lỗi thực thi: " + error.toString() };
    }
}

// -------------------------------------------------------------
// HÀM KHỚP TỰ ĐỘNG CHO 1 HÀNG BÁO "SAI"
// -------------------------------------------------------------
function fixSingleMismatchRow(item, targetUrl, customSheetName) {
    try {
        if (!item || !item.targetRow) {
            return { success: false, message: "Thông tin hàng lỗi không hợp lệ!" };
        }

        var targetFileId = extractSpreadsheetId(targetUrl) || "1zURB1fThVmD1bNY0o6XiDUMsTjLu9IjE4H7sZqF1Ing";
        var targetSs = SpreadsheetApp.openById(targetFileId);
        if (!targetSs) return { success: false, message: "Không mở được File Đích!" };

        var rowIdx = item.targetRow;

        var datesToFix = item.dates || [{
            dateStr: item.dateStr,
            srcVal: item.srcVal,
            srcBg: item.srcBg,
            srcFontColor: item.srcFontColor,
            srcNote: item.srcNote
        }];

        for (var d = 0; d < datesToFix.length; d++) {
            var dObj = datesToFix[d];
            var targetDateStr = dObj.dateStr;
            var dateParts = targetDateStr.split('/');
            var dayNum = parseInt(dateParts[0], 10);
            var monthNum = parseInt(dateParts[1], 10);

            var targetSheet = null;
            if (customSheetName && String(customSheetName).trim() !== "") {
                targetSheet = targetSs.getSheetByName(String(customSheetName).trim());
            }
            if (!targetSheet) {
                targetSheet = targetSs.getSheetByName("CHUYỂN CẦN THÁNG " + monthNum) ||
                    targetSs.getSheetByName("CHUYÊN CẦN THÁNG " + monthNum) ||
                    targetSs.getActiveSheet();
            }

            function findDateColInSheet(sh) {
                try {
                    var lastCol = sh.getLastColumn();
                    if (lastCol < 1) return -1;
                    var header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
                    for (var k = 0; k < header.length; k++) {
                        var cellVal = header[k];
                        if (cellVal instanceof Date) {
                            if (cellVal.getDate() === dayNum && (cellVal.getMonth() + 1) === monthNum) return k + 1;
                        } else if (cellVal) {
                            var strVal = String(cellVal).trim();
                            if (strVal === dayNum + "/" + monthNum || strVal === dateParts[0] + "/" + dateParts[1] || strVal === targetDateStr) return k + 1;
                        }
                    }
                } catch (err) { }
                return -1;
            }

            var tgtDateColIndex = findDateColInSheet(targetSheet);
            if (tgtDateColIndex === -1) {
                var allSheets = targetSs.getSheets();
                for (var s = 0; s < allSheets.length; s++) {
                    var colIdx = findDateColInSheet(allSheets[s]);
                    if (colIdx !== -1) {
                        targetSheet = allSheets[s];
                        tgtDateColIndex = colIdx;
                        break;
                    }
                }
            }

            if (tgtDateColIndex !== -1 && targetSheet) {
                // Chỉ cập nhật Mã Thẻ (Cột D - 4) & Họ Tên (Cột E - 5) nếu có Mã Thẻ hợp lệ ở Nguồn (không phải chuỗi 'Thiếu ở ...')
                if (item.srcId && String(item.srcId).indexOf("Thiếu ở") !== 0) {
                    targetSheet.getRange(rowIdx, 4).setValue(item.srcId).setFontSize(15).setFontWeight("bold").setFontFamily("Times New Roman");
                    targetSheet.getRange(rowIdx, 5).setValue(item.srcName).setFontSize(15).setFontWeight("bold").setFontFamily("Times New Roman");
                }

                // Ghi đè ô ngày: chữ "Sai" -> dữ liệu điểm danh thật ("V")
                var dateCell = targetSheet.getRange(rowIdx, tgtDateColIndex);
                dateCell.setValue(dObj.srcVal || "V");
                dateCell.setBackground(dObj.srcBg || "#dcfce7");
                dateCell.setFontColor(dObj.srcFontColor || "#000000");
                dateCell.setNote(dObj.srcNote || "");
            }
        }

        return {
            success: true,
            message: "✅ Đã khớp thành công cho mã thẻ " + item.srcId + " (" + item.srcName + ") vào Hàng " + rowIdx + "!"
        };

    } catch (err) {
        return { success: false, message: "Lỗi thực thi khớp dòng: " + err.toString() };
    }
}

// -------------------------------------------------------------
// HÀM KHỚP TỰ ĐỘNG TẤT CẢ CÁC HÀNG BÁO "SAI"
// -------------------------------------------------------------
function fixMismatchRows(mismatchList, targetUrl, customSheetName) {
    try {
        if (!mismatchList || !Array.isArray(mismatchList) || mismatchList.length === 0) {
            return { success: false, message: "Không có danh sách hàng báo Sai nào để khớp!" };
        }

        var fixedCount = 0;
        for (var i = 0; i < mismatchList.length; i++) {
            var res = fixSingleMismatchRow(mismatchList[i], targetUrl, customSheetName);
            if (res.success) fixedCount++;
        }

        return {
            success: true,
            message: "⚡ <b>Đã khớp thành công " + fixedCount + " hàng!</b> Hệ thống đã tự động điền Mã Thẻ, Họ Tên sang các hàng File Đích và chuyển tất cả chữ 'Sai' ở các ngày thành dữ liệu điểm danh chuẩn!"
        };

    } catch (err) {
        return { success: false, message: "Lỗi thực thi khớp dữ liệu: " + err.toString() };
    }
}

// -------------------------------------------------------------
// HÀM LƯU / TẢI CÀI ĐẶT MẶC ĐỊNH ĐỒNG BỘ CHUYÊN CẦN (DÙNG SCRIPT PROPERTIES - CHIA SẺ QUA MỌI MÁY/USER)
// -------------------------------------------------------------

function saveCopyDefaultsToServer(settingsJson) {
    try {
        PropertiesService.getScriptProperties().setProperty('CT_COPY_DEFAULTS', settingsJson);
        return { success: true };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function getCopyDefaultsFromServer() {
    try {
        var val = PropertiesService.getScriptProperties().getProperty('CT_COPY_DEFAULTS');
        return { success: true, data: val || null };
    } catch (e) {
        return { success: false, data: null };
    }
}

// -------------------------------------------------------------
// BỘ HÀM XỬ LÝ DỮ LIỆU DS ĐA CÔNG ĐOẠN (KANBAN / CẬP NHẬT / XÓA)
// -------------------------------------------------------------

function getDcdSheet_() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("DS ĐA CÔNG ĐOẠN") ||
        ss.getSheetByName("ĐA CÔNG ĐOẠN") ||
        ss.getSheetByName("DỮ LIỆU ĐA CÔNG ĐOẠN") ||
        ss.getSheetByName("DS_DA_CONG_DOAN");

    if (!sheet) {
        var sheets = ss.getSheets();
        for (var i = 0; i < sheets.length; i++) {
            var sName = sheets[i].getName().toUpperCase();
            if (sName.indexOf("CÔNG ĐOẠN") !== -1 || sName.indexOf("DCD") !== -1 || sName.indexOf("ĐA CÔNG") !== -1) {
                sheet = sheets[i];
                break;
            }
        }
    }
    return sheet;
}

function getKanbanDcdData() {
    try {
        var sheet = getDcdSheet_();
        if (!sheet) {
            return { success: false, message: "Không tìm thấy Sheet 'DS ĐA CÔNG ĐOẠN' trong hệ thống!", data: [] };
        }

        var lastRow = sheet.getLastRow();
        var lastCol = sheet.getLastColumn();

        if (lastRow < 2) {
            return { success: true, data: [] };
        }

        var fetchCols = Math.max(lastCol, 12);
        var range = sheet.getRange(2, 1, lastRow - 1, fetchCols);
        var rawValues = range.getDisplayValues();
        var resultData = [];

        for (var r = 0; r < rawValues.length; r++) {
            var row = rawValues[r];
            var isEmpty = true;
            for (var c = 0; c < row.length; c++) {
                if (row[c] && String(row[c]).trim() !== "") {
                    isEmpty = false;
                    break;
                }
            }
            if (!isEmpty) {
                var rowItem = row.slice(0, 13); // Lấy 13 cột (A -> M)
                rowItem.push(r + 2); // Đẩy rowIndex vào cuối mảng (sẽ là r[13])
                resultData.push(rowItem);
            }
        }

        return {
            success: true,
            data: resultData
        };

    } catch (err) {
        return {
            success: false,
            message: "Lỗi lấy dữ liệu Đa Công Đoạn: " + err.toString(),
            data: []
        };
    }
}



function updateDcdRecord(sheetRowIndex, record) {
    try {
        var sheet = getDcdSheet_();
        if (!sheet) return { success: false, message: "Không tìm thấy Sheet 'DS ĐA CÔNG ĐOẠN'!" };
        if (!sheetRowIndex || sheetRowIndex < 2) return { success: false, message: "Vị trí hàng không hợp lệ!" };

        var rowIdx = parseInt(sheetRowIndex, 10);
        var range = sheet.getRange(rowIdx, 1, 1, 12);
        range.setValues([record]);
        SpreadsheetApp.flush();

        return { success: true, message: "Đã cập nhật hàng thành công!" };
    } catch (err) {
        return { success: false, message: "Lỗi cập nhật hàng: " + err.toString() };
    }
}

function deleteDcdSingleRowBySheetIndex(sheetRowIndex) {
    try {
        var sheet = getDcdSheet_();
        if (!sheet) return { success: false, message: "Không tìm thấy Sheet 'DS ĐA CÔNG ĐOẠN'!" };
        if (!sheetRowIndex || sheetRowIndex < 2) return { success: false, message: "Vị trí hàng không hợp lệ!" };

        var rowIdx = parseInt(sheetRowIndex, 10);
        sheet.deleteRow(rowIdx);
        SpreadsheetApp.flush();

        return { success: true, message: "Đã xóa hàng thành công!" };
    } catch (err) {
        return { success: false, message: "Lỗi xóa hàng: " + err.toString() };
    }
}

function deleteDcdRecordByMaThe(maTheVal) {
    try {
        var sheet = getDcdSheet_();
        if (!sheet) return { success: false, message: "Không tìm thấy Sheet 'DS ĐA CÔNG ĐOẠN'!" };
        if (!maTheVal) return { success: false, message: "Mã Thẻ không hợp lệ!" };

        var searchCode = String(maTheVal).trim().toUpperCase();
        var lastRow = sheet.getLastRow();
        if (lastRow < 2) return { success: true, count: 0 };

        var values = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
        var count = 0;

        for (var r = values.length - 1; r >= 0; r--) {
            if (String(values[r][0]).trim().toUpperCase() === searchCode) {
                sheet.deleteRow(r + 2);
                count++;
            }
        }
        SpreadsheetApp.flush();

        return { success: true, count: count, message: "Đã xóa " + count + " hàng của Mã Thẻ " + maTheVal };
    } catch (err) {
        return { success: false, message: "Lỗi xóa dữ liệu Mã Thẻ: " + err.toString() };
    }
}

// HÀM XỬ LÝ TẢI EXCEL (Dành cho ô số 1)
function processExcel(param1, param2) {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var targetSheet = ss.getSheetByName('出勤数据');
        if (!targetSheet) {
            targetSheet = ss.insertSheet('出勤数据');
        }

        var data;
        if (Array.isArray(param1)) {
            data = param1;
        } else if (typeof param1 === 'string') {
            var metadata = {
                title: param2 || "temp_excel",
                mimeType: MimeType.GOOGLE_SHEETS
            };

            var excelMimeType = param2.toLowerCase().indexOf('.xlsx') > -1 ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" : MimeType.MICROSOFT_EXCEL;

            var boundary = "xxxxx_boundary_xxxxx";
            var requestBody =
                "--" + boundary + "\r\n" +
                "Content-Type: application/json; charset=UTF-8\r\n\r\n" +
                JSON.stringify(metadata) + "\r\n" +
                "--" + boundary + "\r\n" +
                "Content-Type: " + excelMimeType + "\r\n" +
                "Content-Transfer-Encoding: base64\r\n\r\n" +
                param1 + "\r\n" +
                "--" + boundary + "--";

            var response = UrlFetchApp.fetch("https://www.googleapis.com/upload/drive/v2/files?uploadType=multipart&convert=true", {
                method: "post",
                contentType: "multipart/related; boundary=" + boundary,
                headers: {
                    Authorization: "Bearer " + ScriptApp.getOAuthToken()
                },
                payload: requestBody,
                muteHttpExceptions: true
            });

            var resJson = JSON.parse(response.getContentText());
            if (!resJson.id) {
                throw new Error("Không thể chuyển đổi file Excel: " + (resJson.error ? resJson.error.message : response.getContentText()));
            }

            var tempSs = SpreadsheetApp.openById(resJson.id);
            var tempSheet = tempSs.getSheets()[0];
            data = tempSheet.getDataRange().getValues();

            try { DriveApp.getFileById(resJson.id).setTrashed(true); } catch (ez) { }
        }

        if (data && data.length > 0) {
            // Chỉ lấy dữ liệu từ cột B đến H (index 1 đến 7) VÀ cột O (index 14)
            data = data.map(function (row) {
                var newRow = row.slice(1, 8); // B đến H
                newRow.push(row[14] || ""); // Cột O (Ban/Ca)
                return newRow;
            });

            // Xóa dữ liệu cũ trong sheet trước khi dán đè
            targetSheet.clearContents();

            // Dán dữ liệu mới bắt đầu từ dòng 1
            var targetRange = targetSheet.getRange(1, 1, data.length, data[0].length);
            targetRange.setValues(data);

            // Định dạng giờ cho cột F và G (tương ứng cột 6 và 7)
            targetSheet.getRange(1, 6, data.length, 2).setNumberFormat("HH:mm:ss");

            // Bắt buộc đẩy thay đổi xuống Sheet ngay lập tức
            SpreadsheetApp.flush();

            return "Đã tải lên và dán đè " + data.length + " dòng dữ liệu thành công!";
        } else {
            return "Dữ liệu trống!";
        }
    } catch (e) {
        throw new Error("Lỗi xử lý file Excel: " + e.toString());
    }
}

function getAvailableDatesInSheet() {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var targetSheet = ss.getSheetByName('出勤数据');
        if (!targetSheet) return { success: false, message: "Không tìm thấy sheet." };

        var lastRow = targetSheet.getLastRow();
        if (lastRow < 2) return { success: true, dates: [] };

        var dateValues = targetSheet.getRange(2, 5, lastRow - 1, 1).getDisplayValues();
        var uniqueDates = {};
        for (var i = 0; i < dateValues.length; i++) {
            var d = String(dateValues[i][0]).trim();
            if (d && d !== "") {
                uniqueDates[d] = true;
            }
        }

        var datesArray = Object.keys(uniqueDates);
        datesArray.sort(function (a, b) {
            var pA = a.split('/');
            var pB = b.split('/');
            if (pA.length === 3 && pB.length === 3) {
                return new Date(pA[2], pA[1] - 1, pA[0]) - new Date(pB[2], pB[1] - 1, pB[0]);
            }
            return a.localeCompare(b);
        });

        return { success: true, dates: datesArray };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

// ==========================================
// TÍNH NĂNG THÔNG TIN ĐỊA CHỈ
// ==========================================
function setTtdcPasswordAPI(pass) {
    var props = PropertiesService.getScriptProperties();
    if (!pass || pass.trim() === "") {
        props.deleteProperty('TTDC_PASSWORD');
    } else {
        props.setProperty('TTDC_PASSWORD', pass.trim());
    }
    return true;
}

function hasTtdcPasswordAPI() {
    var props = PropertiesService.getScriptProperties();
    var pass = props.getProperty('TTDC_PASSWORD');
    return !!(pass && pass.trim() !== "");
}

function verifyTtdcPasswordAPI(pass) {
    var props = PropertiesService.getScriptProperties();
    var realPass = props.getProperty('TTDC_PASSWORD');
    if (!realPass || realPass.trim() === "") return true;
    return realPass.trim() === pass.trim();
}

function getOrCreateTtdcSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("THÔNG TIN ĐỊA CHỈ");
    if (!sheet) {
        sheet = ss.insertSheet("THÔNG TIN ĐỊA CHỈ");
        var headers = ["Mã Thẻ", "Họ Tên", "Bộ Phận", "Ngày Vào", "Ngày Sinh", "Số Điện Thoại", "Địa Chỉ", "Tình Trạng Hôn Nhân", "Ảnh Base64", "Chức Vụ"];
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e2e8f0");
        sheet.setFrozenRows(1);
    } else {
        // Kiem tra cot Chuc Vu
        var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn() || 1).getValues()[0];
        if (headers.indexOf("Chức Vụ") === -1) {
            var newCol = headers.length + 1;
            sheet.getRange(1, newCol).setValue("Chức Vụ").setFontWeight("bold").setBackground("#e2e8f0");
        }
    }
    return sheet;
}

function getEmployeeInfoForTtdcAPI(maThe) {
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();

        // 1. Kiểm tra trong THÔNG TIN ĐỊA CHỈ
        var sheetTtdc = ss.getSheetByName("THÔNG TIN ĐỊA CHỈ");
        if (sheetTtdc) {
            var valTtdc = sheetTtdc.getDataRange().getDisplayValues();
            var headers = valTtdc[0];
            var cvIdx = headers.indexOf("Chức Vụ");
            for (var i = 1; i < valTtdc.length; i++) {
                if (String(valTtdc[i][0]).trim().toUpperCase() === String(maThe).trim().toUpperCase()) {
                    return {
                        success: true,
                        isExist: true,
                        data: {
                            maThe: valTtdc[i][0],
                            hoTen: valTtdc[i][1],
                            boPhan: valTtdc[i][2],
                            ngayVao: valTtdc[i][3],
                            ngaySinh: valTtdc[i][4],
                            sdt: valTtdc[i][5],
                            diaChi: valTtdc[i][6],
                            honNhan: valTtdc[i][7],
                            anhBase64: valTtdc[i][8],
                            chucVu: (cvIdx > -1) ? valTtdc[i][cvIdx] : ''
                        }
                    };
                }
            }
        }

        // 2. Nếu không có, tìm trong 出勤数据 để lấy tên và ngày vào
        var sheetChuyenCan = ss.getSheetByName("出勤数据");
        if (sheetChuyenCan) {
            var valCc = sheetChuyenCan.getDataRange().getDisplayValues();
            for (var j = 1; j < valCc.length; j++) {
                if (String(valCc[j][1]).trim().toUpperCase() === String(maThe).trim().toUpperCase()) {
                    return {
                        success: true,
                        isExist: false,
                        autoFillData: {
                            hoTen: valCc[j][2],
                            ngayVao: valCc[j][3]
                        }
                    };
                }
            }
        }

        return { success: true, isExist: false };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function saveTtdcAPI(data) {
    try {
        var sheet = getOrCreateTtdcSheet();
        var dataRange = sheet.getDataRange();
        var values = dataRange.getValues();
        var headers = values[0];
        var cvIdx = headers.indexOf("Chức Vụ");
        var rowIndex = -1;

        for (var i = 1; i < values.length; i++) {
            if (String(values[i][0]).trim().toUpperCase() === String(data.maThe).trim().toUpperCase()) {
                rowIndex = i + 1;
                break;
            }
        }

        var rowData = [
            data.maThe,
            data.hoTen,
            data.boPhan,
            (data.ngayVao ? "'" + data.ngayVao : ""),
            (data.ngaySinh ? "'" + data.ngaySinh : ""),
            (data.sdt ? "'" + data.sdt : ""),
            data.diaChi,
            data.honNhan,
            data.anhBase64
        ];

        if (rowIndex > -1) {
            sheet.getRange(rowIndex, 1, 1, 9).setValues([rowData]);
            if (cvIdx > -1) sheet.getRange(rowIndex, cvIdx + 1).setValue(data.chucVu);
            return { success: true, message: "Đã cập nhật thông tin thành công!" };
        } else {
            if (cvIdx > -1) {
                while (rowData.length < cvIdx) rowData.push("");
                rowData[cvIdx] = data.chucVu;
            } else {
                rowData.push(data.chucVu);
            }
            sheet.appendRow(rowData);
            return { success: true, message: "Đã thêm mới thông tin thành công!" };
        }
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function searchTtdcAPI(maThe) {
    try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("THÔNG TIN ĐỊA CHỈ");
        if (!sheet) return { success: false, message: "Chưa có dữ liệu nào trong hệ thống!" };

        var values = sheet.getDataRange().getDisplayValues();
        var headers = values[0];
        var cvIdx = headers.indexOf("Chức Vụ");
        for (var i = 1; i < values.length; i++) {
            if (String(values[i][0]).trim().toUpperCase() === String(maThe).trim().toUpperCase()) {
                var d = {
                    maThe: values[i][0],
                    hoTen: values[i][1],
                    boPhan: values[i][2],
                    ngayVao: values[i][3],
                    ngaySinh: values[i][4],
                    sdt: values[i][5],
                    diaChi: values[i][6],
                    honNhan: values[i][7],
                    anhBase64: values[i][8],
                    chucVu: (cvIdx > -1) ? values[i][cvIdx] : ''
                };
                return { success: true, data: d };
            }
        }
        return { success: false, message: "Không tìm thấy mã thẻ: " + maThe };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function deleteTtdcAPI(maThe) {
    try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("THÔNG TIN ĐỊA CHỈ");
        if (!sheet) return { success: false, message: "Chưa có sheet dữ liệu!" };

        var values = sheet.getDataRange().getValues();
        for (var i = 1; i < values.length; i++) {
            if (String(values[i][0]).trim().toUpperCase() === String(maThe).trim().toUpperCase()) {
                sheet.deleteRow(i + 1);
                return { success: true, message: "Đã xóa thành công mã thẻ " + maThe };
            }
        }
        return { success: false, message: "Không tìm thấy mã thẻ để xóa!" };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function searchTtdcAdvancedAPI(query) {
    try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("THÔNG TIN ĐỊA CHỈ");
        if (!sheet) return { success: false, message: "Chưa có dữ liệu nào trong hệ thống!" };

        var values = sheet.getDataRange().getDisplayValues();
        var headers = values[0];
        var cvIdx = headers.indexOf("Chức Vụ");
        var results = [];

        var qTinh = (query.tinh || "").toLowerCase();
        var qHuyen = (query.huyen || "").toLowerCase();
        var qXa = (query.xa || "").toLowerCase();
        var qDob = (query.dob || "").toLowerCase();

        for (var i = 1; i < values.length; i++) {
            var diaChi = String(values[i][6]).toLowerCase();
            var dob = String(values[i][4]).toLowerCase();

            var match = true;
            if (qTinh && diaChi.indexOf(qTinh) === -1) match = false;
            if (qHuyen && diaChi.indexOf(qHuyen) === -1) match = false;
            if (qXa && diaChi.indexOf(qXa) === -1) match = false;
            if (qDob && dob !== qDob) match = false;

            if (match) {
                results.push({
                    maThe: values[i][0],
                    hoTen: values[i][1],
                    boPhan: values[i][2],
                    ngayVao: values[i][3],
                    ngaySinh: values[i][4],
                    sdt: values[i][5],
                    diaChi: values[i][6],
                    honNhan: values[i][7],
                    anhBase64: values[i][8],
                    chucVu: (cvIdx > -1) ? values[i][cvIdx] : ''
                });
            }
        }

        if (results.length === 0) {
            return { success: false, message: "Không tìm thấy dữ liệu phù hợp!" };
        }

        return { success: true, data: results };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function getTtdcFilterOptionsAPI() {
    try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("THÔNG TIN ĐỊA CHỈ");
        if (!sheet) return { success: false, message: "Chưa có dữ liệu!" };
        var values = sheet.getDataRange().getDisplayValues();

        var tree = {};
        var dobs = {};

        for (var i = 1; i < values.length; i++) {
            var dob = String(values[i][4]).trim();
            if (dob) dobs[dob] = true;

            var addr = String(values[i][6]).trim();
            if (addr) {
                var parts = addr.split(',').map(function (s) { return s.trim(); });
                // Giả định địa chỉ lưu dạng "chitiet, Xã, Huyện, Tỉnh"
                var tinh = parts.length >= 1 ? parts[parts.length - 1] : "";
                var huyen = parts.length >= 2 ? parts[parts.length - 2] : "";
                var xa = parts.length >= 3 ? parts[parts.length - 3] : "";

                if (tinh) {
                    if (!tree[tinh]) tree[tinh] = {};
                    if (huyen) {
                        if (!tree[tinh][huyen]) tree[tinh][huyen] = {};
                        if (xa) {
                            tree[tinh][huyen][xa] = true;
                        }
                    }
                }
            }
        }

        var formattedTree = {};
        for (var t in tree) {
            formattedTree[t] = {};
            for (var h in tree[t]) {
                formattedTree[t][h] = Object.keys(tree[t][h]).sort();
            }
        }

        return {
            success: true,
            tree: formattedTree,
            dobs: Object.keys(dobs).sort()
        };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}
