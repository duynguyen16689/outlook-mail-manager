# Outlook & Hotmail OAuth2 Mail Manager

Hệ thống quản lý tập trung hàng loạt hộp thư **Outlook / Hotmail** thông qua **Microsoft Graph REST API v1.0** và **IMAP XOAUTH2 (`outlook.office365.com:993`)**, hỗ trợ đọc cả **`INBOX`** và **`Junk Email` (Spam)**, tự động nhận diện mã **OTP**, bắt OAuth2 Refresh Token tự động qua **BitBrowser / Puppeteer**, và **Lập lịch tự động quét & Forward thư mới (24h)** về Gmail quản trị (`ip.kdpmanagement+KDP-XXX@gmail.com`).

---

## ✨ Tính năng chính

1. **Đọc đồng thời `INBOX` + `Junk Email` qua API**:
   - Ưu tiên sử dụng **Microsoft Graph REST API v1.0** (`/v1.0/me/mailFolders/inbox/messages` và `/v1.0/me/mailFolders/junkemail/messages`) — đọc được mail ngay cả khi tài khoản bị Microsoft tắt POP/IMAP.
   - Tự động fallback sang **IMAP XOAUTH2 (`outlook.office365.com:993`)** với `ImapFlow` nếu token chỉ có scope IMAP.
   - Sử dụng Public Client ID chính thức của **Mozilla Thunderbird** (`9e5f94bc-e8a4-4e73-b8be-63364c29d753`).
2. **Tự động bóc tách OTP & Cảnh báo KDP**:
   - Nhận diện mã OTP 4–8 chữ số từ tiêu đề và nội dung thư.
   - Phân loại thư quan trọng (`OTP`, `KDP_ALERT`, `GENERAL`).
3. **Bắt OAuth2 Refresh Token tự động (Auto-Browser & BitBrowser)**:
   - Tự động mở đúng profile **BitBrowser** theo mã tài khoản (`KDP-XXX`) để tận dụng cookie/fingerprint tin cậy, vượt xác minh email khôi phục và lấy `refresh_token` tự động.
4. **Lập lịch quét định kỳ (2h/lần) & Auto-Forward 24h**:
   - Tự động quét toàn bộ danh sách hộp thư theo chu kỳ (Interval hoặc Custom khung giờ).
   - Chuyển tiếp nguyên bản (**Native Graph Forward** `/me/messages/{id}/forward`) mọi thư mới trong 24h về `ip.kdpmanagement+KDP-XXX@gmail.com`, chống gửi trùng bằng bảng `OutlookForwardedEmail`.
5. **Hỗ trợ HTTP Proxy riêng cho từng hộp thư**:
   - Tự động đi qua HTTP CONNECT Proxy (`host:port:user:pass`) cho cả OAuth2, Graph API và IMAP, tự động fallback kết nối trực tiếp khi proxy hết hạn.

---

## 📂 Cấu trúc thư mục

```text
src/
├── lib/
│   ├── outlook/
│   │   ├── outlook-oauth.ts         # Đổi Refresh Token -> Access Token (Graph API + IMAP scopes)
│   │   ├── outlook-imap.ts          # Đọc INBOX & Junk qua Microsoft Graph API + IMAP XOAUTH2
│   │   ├── outlook-forwarder.ts     # Native Forward thư mới về Gmail kèm định danh +KDP-XXX
│   │   ├── outlook-scheduler.ts     # Bộ lập lịch chạy ngầm (2h/lần) & quản lý phiên quét
│   │   ├── outlook-auto-browser.ts  # Tự động đăng nhập & lấy Refresh Token qua BitBrowser / Puppeteer
│   │   ├── outlook-rule-setup.ts    # Tự động tạo Rule Forward trên giao diện Outlook Web
│   │   ├── outlook-proxy.ts         # HTTPS over HTTP CONNECT Proxy tunnel
│   │   └── outlook-db.ts            # SQLite Repository (OutlookMailbox & OutlookForwardedEmail)
│   ├── bitbrowser/
│   │   └── bitbrowser-client.ts     # Kết nối BitBrowser Local API (127.0.0.1:54345)
│   └── db/
│       └── prisma.ts                # Khởi tạo Prisma SQLite Client
└── app/
    ├── outlook/
    │   └── page.tsx                 # Giao diện Web quản lý hộp thư, đọc mail & cài đặt lịch quét
    └── api/outlook/
        ├── accounts/route.ts        # API thêm/xóa & Bulk Import (email|pass|refresh_token|client_id)
        ├── check/route.ts           # API đọc hộp thư (INBOX + Junk) & Auto-Forward 24h
        ├── auth/route.ts            # API lấy OAuth2 Token thủ công & tự động
        ├── schedule/route.ts        # API cấu hình & điều khiển bộ lập lịch quét tự động
        └── rules/route.ts           # API thiết lập Rule Forward trên Outlook Web
```

---

## 🚀 Cài đặt & Chạy dự án

```bash
# 1. Cài đặt dependencies
npm install

# 2. Tạo file cấu hình môi trường
cp .env.example .env

# 3. Khởi tạo SQLite Database
npx prisma db push

# 4. Chạy server development (mặc định tại http://localhost:3006/outlook)
npm run dev
```
