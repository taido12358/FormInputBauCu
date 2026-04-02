# Ứng Dụng Nhập Kết Quả Bầu Cử - Phường Sơn Tây

## Yêu Cầu Hệ Thống
- Windows 10/11
- .NET Framework 4.8 hoặc .NET 6+ với Windows Desktop Runtime
- Visual Studio 2019/2022 (để build)

## Cách Chạy

### Bước 1: Mở project
1. Mở file `BauCuSonTay.sln` bằng Visual Studio
2. Visual Studio sẽ tự động restore NuGet packages (NPOI)

### Bước 2: Build & Run
- Nhấn `F5` hoặc `Ctrl+F5` để chạy

## Hướng Dẫn Sử Dụng

### 1. Mở File Excel
- Click **"Mở File Excel"** và chọn file `TH_bau_cu_phuong.xls`

### 2. Chọn Sheet (Đơn Vị Bầu Cử)
- Dropdown **"Đơn vị bầu cử"** sẽ hiển thị 5 sheet: ĐV1, ĐV2, ĐV3, ĐV4, ĐV5

### 3. Chọn Khu Vực Bỏ Phiếu
- Dropdown **"Khu vực bỏ phiếu"** sẽ hiển thị danh sách các khu vực trong sheet đã chọn

### 4. Nhập Dữ Liệu
Các trường bắt buộc:
- Tổng số cử tri (tự động từ Excel nếu có)
- Số lượng cử tri đã tham gia
- Số phiếu phát ra, thu vào, hợp lệ, không hợp lệ
- Số phiếu bầu 1-5 đại biểu
- Số phiếu bầu cho từng ứng viên

Các trường **tự động tính** (màu xanh):
- Tỷ lệ % cử tri tham gia
- Tỷ lệ % phiếu thu vào / phát ra
- Tỷ lệ % phiếu hợp lệ / thu vào
- Tỷ lệ % phiếu không hợp lệ / thu vào

### 5. Kiểm Tra Điều Kiện
- Ô **"Kiểm Tra Điều Kiện"** sẽ hiển thị màu đỏ nếu vi phạm
- Nút **"Lưu vào Excel"** chỉ kích hoạt khi tất cả điều kiện hợp lệ

### 6. Lưu File
- Click **"Lưu vào Excel"** → dữ liệu sẽ được ghi đúng vào hàng tương ứng trong sheet

## Điều Kiện Kiểm Tra

1. Tất cả các trường không được để trống
2. Tỷ lệ % không vượt quá 100%
3. Số cử tri tham gia ≤ Tổng số cử tri
4. Phiếu phát ra ≤ Số cử tri tham gia
5. Phiếu phát ra ≥ Phiếu thu vào
6. Phiếu thu vào = Hợp lệ + Không hợp lệ
7. Phiếu không hợp lệ = Thu vào - Hợp lệ
8. Phiếu hợp lệ = Tổng (1+2+3+4+5 đại biểu)
9. Kiểm tra: 1×PB1 + 2×PB2 + 3×PB3 + 4×PB4 + 5×PB5 = Tổng phiếu ứng viên
10. Phiếu mỗi ứng viên ≤ Phiếu hợp lệ

## Cấu Trúc Dữ Liệu Excel

| Cột | Nội Dung |
|-----|----------|
| B   | Khu vực bỏ phiếu |
| C   | Thuộc tổ dân phố |
| D   | Tổng số khu vực |
| E   | Tổng số cử tri |
| F   | Số lượng cử tri tham gia |
| G   | Tỷ lệ % cử tri (tự động) |
| H   | Số phiếu phát ra |
| I   | Số phiếu thu vào |
| J   | Tỷ lệ % thu vào (tự động) |
| K   | Số phiếu hợp lệ |
| L   | Tỷ lệ % hợp lệ (tự động) |
| M   | Số phiếu không hợp lệ |
| N   | Tỷ lệ % không hợp lệ (tự động) |
| O   | Phiếu bầu 1 đại biểu |
| P   | Phiếu bầu 2 đại biểu |
| Q   | Phiếu bầu 3 đại biểu |
| R   | Phiếu bầu 4 đại biểu |
| S   | Phiếu bầu 5 đại biểu |
| T-AA| Phiếu bầu từng ứng viên |
