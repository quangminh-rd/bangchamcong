document.getElementById('export-excel_phieucungung').addEventListener('click', async function () {
    // Tải template Excel từ server
    try {
        const response = await fetch('./template_phieucungung.xlsx');
        if (!response.ok) throw new Error('Không thể tải template.');
        const buffer = await response.arrayBuffer();

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);
        const worksheet = workbook.getWorksheet(1); // Chọn sheet đầu tiên

        if (!orderDetails) {
            alert('Thông tin đơn hàng chưa được tải.');
            return;
        }
        console.log('orderDetails trước khi xuất Excel:', orderDetails);

        const maHienThi = orderDetails.maHopdong && orderDetails.maHopdong.trim() !== ""
            ? orderDetails.maHopdong
            : orderDetails.maDonhang;

        worksheet.getCell('A3').value = `Số: ${maHienThi || ''}`;
        if (orderDetails.donviPhutrach === "BP. BH1" && orderDetails.phuongThucban !== "Bán chéo") {
            worksheet.getCell('A4').value = `Hôm nay, ngày ${orderDetails.ngayPhatHanh || ''} chúng tôi gồm:`;
            worksheet.getCell('A5').value = `Bên yêu cầu cung cứng (bên A): ${orderDetails.tenNguoilienhe || ''}`;
            worksheet.getCell('H5').value = `Mã khách hàng: ${orderDetails.maKhachHang || ''}`;
            worksheet.getCell('A6').value = `Địa chỉ: ${orderDetails.diachiChitiet || ''}`;
            worksheet.getCell('A8').value = `Điện thoại: ${orderDetails.sdtKhachhang || ''}`;
            worksheet.getCell('A9').value = `Email:  ${orderDetails.emailKhachHang || ''}`;
            worksheet.getCell('A18').value = `Đại diện (Ông/Bà): ${orderDetails.tenNhanvien || ''}`;
            worksheet.getCell('A20').value = `Điện thoại: ${orderDetails.sdtNhanvien || ''}`;

            worksheet.getCell('A538').value = `- Lần 1: Bên A tạm ứng bên B: ${orderDetails.tamUngnpp || '0'} đồng.`;

            worksheet.getCell('J579').value = `${orderDetails.tenNhanvien || ''}`;

        } else if (orderDetails.donviPhutrach === "BP. BH1" && orderDetails.phuongThucban === "Bán chéo") {
            worksheet.getCell('A4').value = `Hôm nay, ngày ${orderDetails.ngayPhatHanh || ''} chúng tôi gồm:`;
            worksheet.getCell('A5').value = `Bên yêu cầu cung cứng (bên A): ${orderDetails.tenKhachhangcuoi || ''}`;
            worksheet.getCell('H5').value = `Mã khách hàng: `;
            worksheet.getCell('A6').value = `Địa chỉ: ${orderDetails.diachiKhachhangcuoi || ''}`;
            worksheet.getCell('A8').value = `Điện thoại: ${orderDetails.sdtKhachhangcuoi || ''}`;
            worksheet.getCell('A9').value = `Email:`;
            worksheet.getCell('A18').value = `Đại diện (Ông/Bà): ${orderDetails.tenNhanvien || ''}`;
            worksheet.getCell('A20').value = `Điện thoại: ${orderDetails.sdtNhanvien || ''}`;

            worksheet.getCell('A538').value = `- Lần 1: Bên A tạm ứng bên B: ${orderDetails.tamUngnpp || '0'} đồng.`;

            worksheet.getCell('J579').value = `${orderDetails.tenNhanvien || ''}`;
        }
        worksheet.getCell('H528').value = orderDetails.tongSobo ? parseFloat(orderDetails.tongSobo) : 0;
        worksheet.getCell('L528').value = orderDetails.congnpp ? parseFloat(formatNumber(orderDetails.congnpp)) : 0;
        worksheet.getCell('H529').value = orderDetails.mucChietkhaunpp ? parseFloat(formatNumber(orderDetails.mucChietkhaunpp)) : 0;
        worksheet.getCell('L529').value = orderDetails.giatriChietkhaunpp ? parseFloat(formatNumber(orderDetails.giatriChietkhaunpp)) : 0;
        worksheet.getCell('L530').value = orderDetails.phiVanchuyenlapdatnpp ? parseFloat(formatNumber(orderDetails.phiVanchuyenlapdatnpp)) : 0;
        worksheet.getCell('H531').value = orderDetails.mucthueGTGTnpp ? parseFloat(formatNumber(orderDetails.mucthueGTGTnpp)) : 0;
        worksheet.getCell('L531').value = orderDetails.thueGTGTnpp ? parseFloat(formatNumber(orderDetails.thueGTGTnpp)) : 0;
        worksheet.getCell('L532').value = orderDetails.tamUngnpp ? parseFloat(formatNumber(orderDetails.tamUngnpp)) : 0;
        worksheet.getCell('L533').value = orderDetails.sotienConthieunpp ? parseFloat(formatNumber(orderDetails.sotienConthieunpp)) : 0;
        worksheet.getCell('A534').value = `Bằng chữ: ${orderDetails.sotienBangchu || ''}`;

        function formatWithCommas(numberString) {
            if (!numberString) return '';
            // Bỏ dấu phân cách hàng nghìn và thay dấu phẩy thập phân bằng dấu chấm
            const num = numberString.replace(/\./g, '').replace(',', '.');
            return num; // Trả về chuỗi số có định dạng chuẩn
        }

        // Điền chi tiết sản phẩm vào Excel
        let startRow = 28; // Ví dụ: bắt đầu từ dòng 28
        orderItems.forEach((item, index) => {
            const row = worksheet.getRow(startRow + index);
            row.getCell(1).value = parseFloat(formatNumber(item.sttTrongdon));
            row.getCell(2).value = item.vitriLapdat;
            row.getCell(3).value = item.maSanphamid;
            row.getCell(4).value = item.diengiai;
            row.getCell(5).value = parseFloat(formatNumber(item.chieuRong || '0')) || '';
            row.getCell(6).value = parseFloat(formatNumber(item.chieuCao));
            row.getCell(7).value = parseFloat(formatWithCommas(item.dienTich));
            row.getCell(8).value = parseFloat(formatNumber(item.soLuong));
            row.getCell(9).value = item.dvt;
            row.getCell(10).value = parseFloat(formatWithCommas(item.khoiLuong));
            row.getCell(11).value = parseFloat(formatNumber(item.dongianpp));
            row.getCell(12).value = parseFloat(formatNumber(item.giabannpp));
        });

        for (let rowNum = 28; rowNum <= 527; rowNum++) {
            const cellValue = worksheet.getCell(`A${rowNum}`).value;

            // Kiểm tra nếu ô A[rowNum] không có dữ liệu hoặc là trống
            if (cellValue === null || cellValue === '') {
                worksheet.getRow(rowNum).hidden = true; // Ẩn dòng tương ứng
            }
        }

        // Kiểm tra và ẩn các dòng từ L529 đến L532 nếu giá trị trong các ô đó là 0 hoặc trống
        if (worksheet.getCell('L529').value === 0 || worksheet.getCell('L529').value === '') {
            worksheet.getRow(529).hidden = true; // Ẩn dòng 529
        }

        if (worksheet.getCell('L530').value === 0 || worksheet.getCell('L530').value === '') {
            worksheet.getRow(530).hidden = true; // Ẩn dòng 530
        }

        if (worksheet.getCell('L531').value === 0 || worksheet.getCell('L531').value === '') {
            worksheet.getRow(531).hidden = true; // Ẩn dòng 531
        }

        if (worksheet.getCell('L532').value === 0 || worksheet.getCell('L532').value === '') {
            worksheet.getRow(532).hidden = true; // Ẩn dòng 532
        }


        // Lưu file Excel và tải về
        const outputBuffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([outputBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `Phiếu cung ứng số ${maHienThi}.xlsx`;

        link.click();
    } catch (error) {
        console.error('Lỗi xuất Excel:', error);
        alert('Không thể xuất Excel. Vui lòng thử lại.');
    }
});