const SPREADSHEET_ID = '10_7GoBzJFFn43um77mBs1NWRzmHqgi_RRutVVR1Nf7g';
const RANGE = 'cham_cong_theo_ca!A:X';
const API_KEY = 'AIzaSyA9g2qFUolpsu3_HVHOebdZb0NXnQgXlFM';

document.addEventListener('DOMContentLoaded', () => {
    const script = document.createElement('script');
    script.src = "https://apis.google.com/js/api.js";
    script.onload = initializeGAPI;
    document.body.appendChild(script);
    const excelBtn = document.getElementById('export-excel');
    if (excelBtn) {
        const excelBtn = document.getElementById('export-excel');
        if (excelBtn) {
            excelBtn.addEventListener('click', async () => {
                const workbook = new ExcelJS.Workbook();
                const worksheet = workbook.addWorksheet('Bảng chấm công');

                // Tạo lưới ảo để xử lý chính xác vị trí ô
                const table = document.querySelector('.bordered-table');
                const rows = Array.from(table.querySelectorAll('tr'));
                const grid = [];
                const merges = [];

                // Khởi tạo grid với các ô trống
                rows.forEach((tr, rowIdx) => {
                    grid[rowIdx] = [];
                });

                // Đổ dữ liệu vào grid và ghi nhận các ô cần merge
                rows.forEach((tr, rowIdx) => {
                    const cells = Array.from(tr.querySelectorAll('th, td'));
                    let colIdx = 0;

                    // Tìm cột trống tiếp theo
                    while (grid[rowIdx][colIdx] !== undefined) {
                        colIdx++;
                    }

                    cells.forEach(cell => {
                        const colspan = parseInt(cell.getAttribute('colspan')) || 1;
                        const rowspan = parseInt(cell.getAttribute('rowspan')) || 1;

                        // Ghi nhận merge
                        if (rowspan > 1 || colspan > 1) {
                            merges.push({
                                start: { row: rowIdx, col: colIdx },
                                end: { row: rowIdx + rowspan - 1, col: colIdx + colspan - 1 }
                            });
                        }

                        // Đánh dấu các ô bị chiếm chỗ
                        for (let r = 0; r < rowspan; r++) {
                            for (let c = 0; c < colspan; c++) {
                                if (r === 0 && c === 0) {
                                    // Ô chính
                                    grid[rowIdx + r][colIdx + c] = {
                                        content: cell.innerText.trim(),
                                        element: cell
                                    };
                                } else {
                                    // Ô bị merge
                                    grid[rowIdx + r][colIdx + c] = 'merged';
                                }
                            }
                        }

                        colIdx += colspan;
                    });
                });

                // Tạo worksheet từ grid
                grid.forEach((row, rowIdx) => {
                    const excelRow = worksheet.addRow([]);

                    row.forEach((cell, colIdx) => {
                        if (cell === 'merged') return; // Bỏ qua ô đã merge

                        const excelCell = excelRow.getCell(colIdx + 1);
                        excelCell.value = cell.content;

                        const isDayCell = cell.element.classList.contains('borderedcol-day');

                        excelCell.font = {
                            name: 'Times New Roman',
                            size: isDayCell ? 8 : 11, // Font size 8 cho ô ngày
                            color: cell.element.classList.contains('highlight-x')
                                ? { argb: 'FFFF0000' } : undefined
                        };

                        // Áp dụng fill cho các ô highlight-cn và header-green
                        if (cell.element.classList.contains('highlight-cn')) {
                            excelCell.fill = {
                                type: 'pattern',
                                pattern: 'solid',
                                fgColor: { argb: 'FF09B9FF' }
                            };
                        }
                        if (cell.element.classList.contains('highlight-header-green')) {
                            excelCell.fill = {
                                type: 'pattern',
                                pattern: 'solid',
                                fgColor: { argb: 'FFB6DB99' }
                            };
                        }

                        // Căn giữa
                        excelCell.alignment = {
                            vertical: 'middle',
                            horizontal: 'center',
                            wrapText: true
                        };

                        // Border
                        excelCell.border = {
                            top: { style: 'thin' },
                            left: { style: 'thin' },
                            bottom: { style: 'thin' },
                            right: { style: 'thin' }
                        };

                    });
                });

                // Áp dụng các merge
                merges.forEach(merge => {
                    worksheet.mergeCells(
                        merge.start.row + 1,
                        merge.start.col + 1,
                        merge.end.row + 1,
                        merge.end.col + 1
                    );
                });

                // Đặt chiều rộng cột
                worksheet.getColumn(1).width = 5;   // STT
                worksheet.getColumn(2).width = 25;  // Tên NV
                worksheet.getColumn(3).width = 10;  // Buổi

                // Bôi đậm dòng 1 và dòng 2 (ExcelJS dùng chỉ số bắt đầu từ 1)
                [1, 2].forEach(rowNumber => {
                    const row = worksheet.getRow(rowNumber);
                    row.eachCell(cell => {
                        cell.font = {
                            ...cell.font,
                            bold: true
                        };
                    });
                });

                // Căn trái cột 2 và cột 3
                [2, 3].forEach(col => {
                    worksheet.getColumn(col).eachCell({ includeEmpty: true }, cell => {
                        cell.alignment = {
                            vertical: 'middle',
                            horizontal: 'left',
                            wrapText: true
                        };
                    });
                });

                // Cột ngày (giả sử có 31 ngày)
                for (let i = 4; i <= 34; i++) {
                    worksheet.getColumn(i).width = 4;
                }

                // Cột tổng (3 cột cuối)
                const totalColumns = 3;
                for (let i = 0; i < totalColumns; i++) {
                    const colIndex = worksheet.columnCount - totalColumns + i + 1;
                    worksheet.getColumn(colIndex).width = 8;
                }

                // Đóng băng header
                worksheet.views = [{
                    state: 'frozen',
                    ySplit: 2 // Số dòng header
                }];

                // Xuất file
                const buffer = await workbook.xlsx.writeBuffer();
                const blob = new Blob([buffer], {
                    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                });

                const a = document.createElement('a');
                a.href = URL.createObjectURL(blob);
                const thangSelect = document.getElementById('thangSelect');
                const namSelect = document.getElementById('namSelect');
                const thang = thangSelect ? thangSelect.value : new Date().getMonth() + 1;
                const nam = namSelect ? namSelect.value : new Date().getFullYear();
                a.download = `Bảng chấm công tháng ${parseInt(thang)} năm ${nam}.xlsx`;
                a.click();
            });
        }
    }
});

function getDataFromURI() {
    const url = window.location.href;
    const match = url.match(/maNhanvien=([^?&]*)/);
    return match ? decodeURIComponent(match[1]) : null;
}

function getKhuVucFromURI() {
    const url = window.location.href;
    const match = url.match(/khuVuc=([^?&]*)/);
    return match ? decodeURIComponent(match[1]).split('-') : [];
}

function updateContent(message) {
    const el = document.getElementById("contentMessage");
    if (el) el.innerText = message;
}

function formatNumber(n) {
    const num = parseFloat(n);
    if (isNaN(num)) return '';
    return num % 1 === 0 ? num.toString() : num.toString().replace(/\.?0+$/, '');
}

function getWeekdayShort(day, month, year) {
    const date = new Date(`${year}-${month}-${day}`);
    const weekday = date.getDay(); // 0: CN, 1: T2, ..., 6: T7
    return ['CN', 'T2', 'T3', 'T4', 'T5', 'T6', 'T7'][weekday];
}

function populateSelect(maList, rows) {
    const nhanvienSelect = document.getElementById('nhanvienSelect');
    const thangSelect = document.getElementById('thangSelect');
    const namSelect = document.getElementById('namSelect');

    // Tạo Map từ mã nhân viên => tên nhân viên
    const maToTenMap = new Map();
    rows.forEach(row => {
        const ma = row[2];
        const ten = row[3];
        if (ma && ten && !maToTenMap.has(ma)) {
            maToTenMap.set(ma, ten);
        }
    });

    const allOpt = document.createElement('option');
    allOpt.value = 'ALL';
    allOpt.textContent = 'Tất cả';
    allOpt.selected = true;
    nhanvienSelect.appendChild(allOpt);

    // Tạo danh sách mã nhân viên
    maList.forEach(ma => {
        const opt = document.createElement('option');
        opt.value = ma;
        opt.textContent = maToTenMap.get(ma) || ma;
        nhanvienSelect.appendChild(opt);
    });

    const khuvucSelect = document.getElementById('khuvucSelect');
    const danhSachKhuVuc = getKhuVucFromURI();

    // Tạo option khu vực
    const allKhuvucOpt = document.createElement('option');
    allKhuvucOpt.value = 'ALL';
    allKhuvucOpt.textContent = 'Tất cả';
    khuvucSelect.appendChild(allKhuvucOpt);

    danhSachKhuVuc.forEach(kv => {
        const opt = document.createElement('option');
        opt.value = kv;
        opt.textContent = kv;
        khuvucSelect.appendChild(opt);
    });
    khuvucSelect.value = 'ALL'; // Default

    // Tạo danh sách tháng (01 - 12)
    for (let i = 1; i <= 12; i++) {
        const thang = String(i).padStart(2, '0');
        const opt = document.createElement('option');
        opt.value = thang;
        opt.textContent = `Tháng ${thang}`;
        thangSelect.appendChild(opt);
    }

    // Chọn mặc định: mã đầu tiên và tháng hiện tại
    const today = new Date();
    const thangHienTai = String(today.getMonth() + 1).padStart(2, '0');
    thangSelect.value = thangHienTai;

    // Lấy năm hiện tại
    const namHientai = new Date().getFullYear();
    // Giả sử hiển thị từ 2022 đến năm hiện tại + 1
    for (let y = namHientai - 2; y <= namHientai + 1; y++) {
        const opt = document.createElement('option');
        opt.value = y;
        opt.textContent = `Năm ${y}`;
        namSelect.appendChild(opt);
    }
    namSelect.value = namHientai;

    // Sự kiện thay đổi
    nhanvienSelect.addEventListener('change', () => {
        const selectedNhanViens = Array.from(nhanvienSelect.selectedOptions).map(opt => opt.value);
        fetchAndRenderFor(khuvucSelect.value, selectedNhanViens, thangSelect.value, namSelect.value, rows);
    });


    // Gắn event cho khu vực
    khuvucSelect.addEventListener('change', () => {
        fetchAndRenderFor(khuvucSelect.value, nhanvienSelect.value, thangSelect.value, namSelect.value, rows);
    });

    namSelect.addEventListener('change', () => {
        fetchAndRenderFor(khuvucSelect.value, nhanvienSelect.value, thangSelect.value, namSelect.value, rows);
    });

    thangSelect.addEventListener('change', () => {
        fetchAndRenderFor(khuvucSelect.value, nhanvienSelect.value, thangSelect.value, namSelect.value, rows);
    });

    // Sau khi setup xong, gọi fetchAndRenderFor tự động
    const maNhanvien = nhanvienSelect.value;
    const thang = thangSelect.value;
    const nam = namSelect.value;
    const khuVuc = khuvucSelect.value;

    fetchAndRenderFor(khuVuc, maNhanvien, thang, nam, rows);

}

async function fetchAndRenderFor(khuVuc, maNhanvien, thang, nam) {
    const selectedKhuVuc = document.getElementById('khuvucSelect').value;
    try {
        const rows = await fetchData();
        const danhSachMa = getDataFromURI().split('-');

        const tableHeader = document.getElementById('tableHeader');
        const tableBody = document.getElementById('tableBody');
        tableHeader.innerHTML = '';
        tableBody.innerHTML = '';

        let allFiltered = [];
        const maList = Array.isArray(maNhanvien)
            ? maNhanvien.includes('ALL') ? danhSachMa : maNhanvien
            : (maNhanvien === 'ALL' ? danhSachMa : [maNhanvien]);


        for (const ma of maList) {
            const filtered = filterDataByNhanVienThangNam(rows, ma, thang, nam, selectedKhuVuc);
            if (filtered.length > 0) allFiltered.push({ ma, data: filtered });
        }

        if (allFiltered.length === 0) {
            updateContent("Không có dữ liệu chấm công!");
            showPopup("Không có dữ liệu chấm công!");
            return;
        }

        // Lấy tháng & năm từ dòng đầu tiên hợp lệ
        const exampleRow = allFiltered[0].data[0];
        const [dd, mm, yyyy] = (exampleRow[5] || '').split('/');
        const daysInMonth = Array.from({ length: new Date(yyyy, mm, 0).getDate() }, (_, i) => String(i + 1).padStart(2, '0'));

        // Dựng header 1 lần
        renderChamCongHeader(daysInMonth, mm, yyyy);

        // Dựng từng dòng dữ liệu
        let stt = 1;
        for (const item of allFiltered) {
            const processed = processChamCongData(item.data);
            const tenNV = item.data[0][3] || item.ma;
            renderChamCongRow(processed, tenNV, stt++, daysInMonth, mm, yyyy);
        }

        updateContent(""); // Xoá thông báo cũ nếu có

    } catch (err) {
        console.error(err);
        updateContent("Lỗi khi tải dữ liệu.");
        showPopup("Lỗi khi tải dữ liệu chấm công!");
    }
}

function initializeGAPI() {
    gapi.load('client', async () => {
        try {
            await gapi.client.init({
                apiKey: API_KEY,
                discoveryDocs: ['https://sheets.googleapis.com/$discovery/rest?version=v4']
            });

            const maNhanviens = getDataFromURI();
            if (!maNhanviens) {
                updateContent("Mời chọn thông tin!");
                showPopup("Mời chọn thông tin!");
                return;
            }


            const danhSachMa = maNhanviens.split('-');
            const rows = await fetchData();
            populateSelect(danhSachMa, rows);

        } catch (error) {
            console.error(error);
            updateContent("Lỗi khởi tạo hoặc tải dữ liệu.");
        }
    });
}


async function fetchData() {
    const res = await gapi.client.sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: RANGE
    });
    return res.result.values || [];
}

function filterDataByNhanVienThangNam(rows, maNhanvien, thang, nam, khuVuc = 'ALL') {
    return rows.filter(row => {
        const ma = row[2];
        const ngay = row[5];
        const kv = row[19]; // Cột T
        const ten = row[3];
        const [dd, mm, yyyy] = (ngay || '').split('/');

        const matchMa = (maNhanvien === 'ALL') || maNhanvien === ma;
        const matchThang = mm === thang;
        const matchNam = yyyy === nam;
        const matchKhuVuc = khuVuc === 'ALL' || kv === khuVuc;

        return matchMa && matchThang && matchNam && matchKhuVuc;
    });
}


function processChamCongData(filtered) {
    const sang = {}, chieu = {};
    let tongCongS = 0, tongCongC = 0;
    let tongTC_S = 0, tongTC_C = 0;
    let diMuonS = 0, diMuonC = 0;
    const daySet = new Set();

    // Lấy tháng & năm từ dòng đầu tiên
    const [dd, mm, yyyy] = (filtered[0][5] || '').split('/');
    const totalDays = new Date(parseInt(yyyy), parseInt(mm), 0).getDate(); // Số ngày trong tháng

    const daysInMonth = Array.from({ length: totalDays }, (_, i) => String(i + 1).padStart(2, '0'));


    filtered.forEach(row => {
        const ngay = row[5];
        const buoi = (row[6] || '').toLowerCase();

        const tangCa = (row[13]).toString().replace(',', '.');
        const diMuon = parseFloat(row[14]) || 0;

        const raw = (row[12] || '').toString().replace(',', '.');
        const val = parseFloat(raw);
        const chamCong = !isNaN(val) ? (val === 4 ? 'V' : (val === 0 ? 'X' : val.toString())) : '';
        const congSo = !isNaN(val) ? val : 0;

        const day = ngay.split('/')[0].padStart(2, '0');
        daySet.add(day);

        if (buoi.includes('sáng')) {
            sang[day] = chamCong;
            tongCongS += congSo / 8;
            tongTC_S += tangCa / 8;
            diMuonS += diMuon;
        } else if (buoi.includes('chiều')) {
            chieu[day] = chamCong;
            tongCongC += congSo / 8;
            tongTC_C += tangCa / 8;
            diMuonC += diMuon;
        }
    });

    const days = daysInMonth;

    return {
        sang, chieu, days: daysInMonth,
        tongCongS, tongCongC,
        tongTC_S, tongTC_C,
        diMuonS, diMuonC,
        currentMonth: mm,
        namHientai: yyyy
    };

}

function renderChamCongTableToIds(data, tenNhanVien, headerId, bodyId) {
    const {
        sang, chieu, days,
        tongCongS, tongCongC,
        tongTC_S, tongTC_C,
        diMuonS, diMuonC,
        currentMonth, currentYear
    } = data;

    const tongCongAll = formatNumber(tongCongS + tongCongC);
    const tongTCAll = formatNumber(tongTC_S + tongTC_C);
    const diMuonAll = diMuonS + diMuonC;

    const header = document.getElementById(headerId);
    let headerHtml = `<tr>
        <th class="borderedcol-1 highlight-header-green" rowspan="2">STT</th>
        <th class="borderedcol-2 highlight-header-green" rowspan="2">Tên nhân viên</th>
        <th class="borderedcol-3 highlight-header-green" rowspan="2">Buổi</th>`;

    const cnColumns = [];
    days.forEach((d, index) => {
        const weekday = getWeekdayShort(d, currentMonth, currentYear);
        if (weekday === 'CN') cnColumns.push(index);
        headerHtml += `<th class="borderedcol-day ${weekday === 'CN' ? 'highlight-cn' : 'highlight-header-green'}">${weekday}</th>`;
    });

    headerHtml += `
        <th class="borderedcol-total highlight-header-green" rowspan="2">Số công</th>
        <th class="borderedcol-total highlight-header-green" rowspan="2">Tăng ca</th>
        <th class="borderedcol-total highlight-header-green" rowspan="2">Đi muộn</th>
        </tr><tr>`;

    days.forEach((d, index) => {
        const isCN = cnColumns.includes(index);
        headerHtml += `<th class="borderedcol-day ${isCN ? 'highlight-cn' : 'highlight-header-green'}">${d}</th>`;
    });
    headerHtml += `</tr>`;

    header.innerHTML = headerHtml;

    const body = document.getElementById(bodyId);
    let bodyHtml = `<tr>
        <td class="borderedcol-1" rowspan="2">1</td>
        <td class="borderedcol-2" rowspan="2">${tenNhanVien}</td>
        <td class="borderedcol-3">Buổi sáng</td>`;

    days.forEach((d, index) => {
        const val = sang.hasOwnProperty(d) ? sang[d] : 'X';
        const classes = [];
        if (val === 'X') classes.push('highlight-x');
        if (cnColumns.includes(index)) classes.push('highlight-cn');
        bodyHtml += `<td class="borderedcol-day ${classes.join(' ')}">${val}</td>`;
    });

    bodyHtml += `
        <td class="borderedcol-total" rowspan="2">${tongCongAll}</td>
        <td class="borderedcol-total" rowspan="2">${tongTCAll}</td>
        <td class="borderedcol-total" rowspan="2">${diMuonAll}</td>
        </tr><tr>
        <td class="borderedcol-3">Buổi chiều</td>`;

    days.forEach((d, index) => {
        const val = chieu.hasOwnProperty(d) ? chieu[d] : 'X';
        const classes = [];
        if (val === 'X') classes.push('highlight-x');
        if (cnColumns.includes(index)) classes.push('highlight-cn');
        bodyHtml += `<td class="borderedcol-day ${classes.join(' ')}">${val}</td>`;
    });

    bodyHtml += `</tr>`;
    body.innerHTML = bodyHtml;
}


function showPopup(message) {
    const popup = document.getElementById('customPopup');
    const messageBox = document.getElementById('popupMessage');
    messageBox.textContent = message || "Thông báo!";
    popup.style.display = 'block';
}

function closePopup() {
    const popup = document.getElementById('customPopup');
    popup.style.display = 'none';
}

// Ẩn popup nếu người dùng click ra ngoài nội dung
window.addEventListener('click', function (event) {
    const modal = document.getElementById('customPopup');
    if (event.target === modal) {
        modal.style.display = 'none';
    }
});

function renderChamCongHeader(days, mm, yyyy) {
    const header = document.getElementById('tableHeader');
    const cnColumns = [];

    // Dòng thứ
    let html = `<tr>
        <th class="borderedcol-1 highlight-header-green" rowspan="2">STT</th>
        <th class="borderedcol-2 highlight-header-green" rowspan="2">Tên nhân viên</th>
        <th class="borderedcol-3 highlight-header-green" rowspan="2">Buổi</th>`;
    days.forEach((d, idx) => {
        const dayLabel = getWeekdayShort(d, mm, yyyy);
        if (dayLabel === 'CN') cnColumns.push(idx);
        html += `<th class="borderedcol-day ${dayLabel === 'CN' ? 'highlight-cn' : 'highlight-header-green'}">${dayLabel}</th>`;
    });
    html += `<th class="borderedcol-total highlight-header-green" rowspan="2">Số công</th>
             <th class="borderedcol-total highlight-header-green" rowspan="2">Tăng ca</th>
             <th class="borderedcol-total highlight-header-green" rowspan="2">Đi muộn</th>
        </tr>`;

    // Dòng ngày
    html += `<tr>`;
    days.forEach((d, idx) => {
        html += `<th class="borderedcol-day ${cnColumns.includes(idx) ? 'highlight-cn' : 'highlight-header-green'}">${d}</th>`;
    });
    html += `</tr>`;

    header.innerHTML = html;
}
function renderChamCongRow(data, tenNhanVien, stt, days, mm, yyyy) {
    const {
        sang, chieu,
        tongCongS, tongCongC,
        tongTC_S, tongTC_C,
        diMuonS, diMuonC
    } = data;

    const cnColumns = days.map((d, i) => getWeekdayShort(d, mm, yyyy) === 'CN' ? i : -1).filter(i => i >= 0);

    const tongCongAll = formatNumber(tongCongS + tongCongC);
    const tongTCAll = formatNumber(tongTC_S + tongTC_C);
    const diMuonAll = diMuonS + diMuonC;

    const body = document.getElementById('tableBody');
    let html = '';

    // Dòng sáng
    html += `<tr class="row-no-bottom">
        <td class="borderedcol-1" rowspan="2">${stt}</td>
        <td class="borderedcol-2" rowspan="2">${tenNhanVien}</td>
        <td class="borderedcol-3">Buổi sáng</td>`;
    days.forEach((d, idx) => {
        const val = sang[d] || 'X';
        const classes = [];
        if (val === 'X') classes.push('highlight-x');
        if (cnColumns.includes(idx)) classes.push('highlight-cn');
        html += `<td class="borderedcol-day ${classes.join(' ')}">${val}</td>`;
    });
    html += `<td class="borderedcol-total" rowspan="2">${tongCongAll}</td>
             <td class="borderedcol-total" rowspan="2">${tongTCAll}</td>
             <td class="borderedcol-total" rowspan="2">${diMuonAll}</td>
         </tr>`;

    // Dòng chiều
    html += `<tr class="row-dashed-middle"><td class="borderedcol-3">Buổi chiều</td>`;
    days.forEach((d, idx) => {
        const val = chieu[d] || 'X';
        const classes = [];
        if (val === 'X') classes.push('highlight-x');
        if (cnColumns.includes(idx)) classes.push('highlight-cn');
        html += `<td class="borderedcol-day ${classes.join(' ')}">${val}</td>`;
    });
    html += `</tr>`;


    body.innerHTML += html;
}