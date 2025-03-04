var DocTienBangChu = function () {
    this.ChuSo = new Array(" không ", " một ", " hai ", " ba ", " bốn ", " năm ", " sáu ", " bảy ", " tám ", " chín ");
    this.Tien = new Array("", " nghìn", " triệu", " tỷ", " nghìn tỷ", " triệu tỷ");
};

DocTienBangChu.prototype.docSo3ChuSo = function (baso) {
    var tram;
    var chuc;
    var donvi;
    var KetQua = "";
    tram = parseInt(baso / 100);
    chuc = parseInt((baso % 100) / 10);
    donvi = baso % 10;
    if (tram == 0 && chuc == 0 && donvi == 0) return "";
    if (tram != 0) {
        KetQua += this.ChuSo[tram] + " trăm ";
        if ((chuc == 0) && (donvi != 0)) KetQua += " linh ";
    }
    if ((chuc != 0) && (chuc != 1)) {
        KetQua += this.ChuSo[chuc] + " mươi";
        if ((chuc == 0) && (donvi != 0)) KetQua = KetQua + " linh ";
    }
    if (chuc == 1) KetQua += " mười ";
    switch (donvi) {
        case 1:
            if ((chuc != 0) && (chuc != 1)) {
                KetQua += " mốt ";
            }
            else {
                KetQua += this.ChuSo[donvi];
            }
            break;
        case 5:
            if (chuc == 0) {
                KetQua += this.ChuSo[donvi];
            }
            else {
                KetQua += " lăm ";
            }
            break;
        default:
            if (donvi != 0) {
                KetQua += this.ChuSo[donvi];
            }
            break;
    }
    return KetQua;
}

DocTienBangChu.prototype.doc = function (SoTien) {
    var lan = 0;
    var i = 0;
    var so = 0;
    var KetQua = "";
    var tmp = "";
    var soAm = false;
    var ViTri = new Array();
    if (SoTien < 0) soAm = true;//return "Số tiền âm !";
    if (SoTien == 0) return "Không đồng";//"Không đồng !";
    if (SoTien > 0) {
        so = SoTien;
    }
    else {
        so = -SoTien;
    }
    if (SoTien > 8999999999999999) {
        //SoTien = 0;
        return "";//"Số quá lớn!";
    }
    ViTri[5] = Math.floor(so / 1000000000000000);
    if (isNaN(ViTri[5]))
        ViTri[5] = "0";
    so = so - parseFloat(ViTri[5].toString()) * 1000000000000000;
    ViTri[4] = Math.floor(so / 1000000000000);
    if (isNaN(ViTri[4]))
        ViTri[4] = "0";
    so = so - parseFloat(ViTri[4].toString()) * 1000000000000;
    ViTri[3] = Math.floor(so / 1000000000);
    if (isNaN(ViTri[3]))
        ViTri[3] = "0";
    so = so - parseFloat(ViTri[3].toString()) * 1000000000;
    ViTri[2] = parseInt(so / 1000000);
    if (isNaN(ViTri[2]))
        ViTri[2] = "0";
    ViTri[1] = parseInt((so % 1000000) / 1000);
    if (isNaN(ViTri[1]))
        ViTri[1] = "0";
    ViTri[0] = parseInt(so % 1000);
    if (isNaN(ViTri[0]))
        ViTri[0] = "0";
    if (ViTri[5] > 0) {
        lan = 5;
    }
    else if (ViTri[4] > 0) {
        lan = 4;
    }
    else if (ViTri[3] > 0) {
        lan = 3;
    }
    else if (ViTri[2] > 0) {
        lan = 2;
    }
    else if (ViTri[1] > 0) {
        lan = 1;
    }
    else {
        lan = 0;
    }
    for (i = lan; i >= 0; i--) {
        tmp = this.docSo3ChuSo(ViTri[i]);
        KetQua += tmp;
        if (ViTri[i] > 0) KetQua += this.Tien[i];
        if ((i > 0) && (tmp.length > 0)) KetQua += '';//',';//&& (!string.IsNullOrEmpty(tmp))
    }
    if (KetQua.substring(KetQua.length - 1) == ',') {
        KetQua = KetQua.substring(0, KetQua.length - 1);
    }
    KetQua = KetQua.substring(1, 2).toUpperCase() + KetQua.substring(2);
    if (soAm) {
        return "Âm " + KetQua + " đồng";//.substring(0, 1);//.toUpperCase();// + KetQua.substring(1);
    }
    else {
        return KetQua + " đồng";//.substring(0, 1);//.toUpperCase();// + KetQua.substring(1);
    }
}

function formatNumber(numberString) {
    if (!numberString) return '';
    // Loại bỏ tất cả dấu chấm
    const num = numberString.replace(/\./g, '');
    const formatted = parseFloat(num).toString();
    return formatted.replace('.', ',');
}

function formatWithCommas(numberString) {
    if (!numberString) return '';
    const num = numberString.replace(',', '.');
    return parseFloat(num).toLocaleString('it-IT');
}

const SPREADSHEET_ID = '1idd6prryF4SFemjPHQpKUwXQCg9rRxsn8DE6DtXCeP4';
const RANGE = 'nhap_xuat_kho!A:N'; // Mở rộng phạm vi đến cột N
const RANGE_CHITIET = 'nhap_xuat_kho_chi_tiet!B:V'; // Dải dữ liệu từ sheet 'nhap_xuat_kho_chi_tiet'
const API_KEY = 'AIzaSyA9g2qFUolpsu3_HVHOebdZb0NXnQgXlFM';

// Lấy giá trị từ URI sau dấu "?" cho các tham số cụ thể
function getDataFromURI() {
    const url = window.location.href;

    // Sử dụng RegEx để trích xuất giá trị của ma_phieu_xuat, nguoi_tao, và nhap_tai_kho
    const maPhieuURIMatch = url.match(/ma_phieu=([^?&]*)/);
    const nguoiTaoMatch = url.match(/nguoi_tao=([^?&]*)/);

    // Gán các giá trị vào các biến
    const maPhieuURI = maPhieuURIMatch ? decodeURIComponent(maPhieuURIMatch[1]) : null;
    const nguoiTaoURI = nguoiTaoMatch ? decodeURIComponent(nguoiTaoMatch[1]) : null;

    // Trả về một đối tượng chứa các giá trị
    return {
        maPhieuURI,
        nguoiTaoURI
    };
}

function extractDay(dateString) {
    if (!dateString) return '';

    // Chuẩn hóa định dạng ngày về "DD/MM/YYYY"
    const parts = dateString.split(/[-/]/); // Chấp nhận cả "-" và "/"
    if (parts.length === 3) {
        let day, month, year;

        if (parts[0].length === 4) {
            // Định dạng ban đầu là "YYYY/MM/DD" hoặc "YYYY-MM-DD"
            [year, month, day] = parts;
        } else if (parts[1].length === 4) {
            // Định dạng ban đầu là "DD/MM/YYYY" (đã đúng)
            [day, month, year] = parts;
        } else {
            // Giả định định dạng "MM/DD/YYYY"
            [month, day, year] = parts;
        }

        // Đảm bảo các phần đều đủ 2 chữ số (nếu cần)
        day = day.padStart(2, '0');
        month = month.padStart(2, '0');

        // Chuẩn hóa thành "DD/MM/YYYY"
        dateString = `${day}/${month}/${year}`;
    }

    // Trích xuất ngày từ định dạng "DD/MM/YYYY"
    return dateString.split('/')[0];
}


// Hàm để tải Google API Client
function loadGapiAndInitialize() {
    const script = document.createElement('script');
    script.src = "https://apis.google.com/js/api.js"; // Đường dẫn đến Google API Client
    script.onload = initialize; // Gọi hàm `initialize` sau khi thư viện được tải xong
    script.onerror = () => console.error('Failed to load Google API Client.');
    document.body.appendChild(script); // Gắn thẻ script vào tài liệu
}

// Hàm khởi tạo sau khi Google API Client được tải
function initialize() {
    gapi.load('client', async () => {
        try {
            await gapi.client.init({
                apiKey: API_KEY,
                discoveryDocs: ['https://sheets.googleapis.com/$discovery/rest?version=v4']
            });

            const uriData = getDataFromURI();
            if (!uriData.maPhieuURI) {
                updateContent('No valid data found in URI.');
                return;
            }

            findRowInSheet(uriData.maPhieuURI);
            findDetailsInSheet(uriData.maPhieuURI);

        } catch (error) {
            updateContent('Initialization error: ' + error.message);
            console.error('Initialization Error:', error);
        }
    });
}

// Gọi hàm tải Google API Client khi DOM đã sẵn sàng
document.addEventListener('DOMContentLoaded', () => {
    loadGapiAndInitialize();
});

function updateContent(message) {
    const contentElement = document.getElementById('content'); // Thay 'content' bằng ID của phần tử HTML cần hiển thị
    if (contentElement) {
        contentElement.textContent = message;
    } else {
        console.warn('Element with ID "content" not found.');
    }
}


// Tìm chỉ số dòng chứa dữ liệu khớp trong cột B và lấy các giá trị từ các cột khác
let orderDetails = null; // Thông tin đơn hàng chính
let orderItems = [];

async function findRowInSheet(maPhieuURI) {
    const uriData = getDataFromURI();

    try {
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: RANGE,
        });

        const rows = response.result.values;
        if (!rows || rows.length === 0) {
            updateContent('No data found.');
            return;
        }

        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];

            const bColumnValue = row[0]; // Cột A
            if (bColumnValue === maPhieuURI) {
                // Lưu dữ liệu vào biến toàn cục
                orderDetails = {
                    maPhieu: row[0] || '', // Cột A
                    loaiPhieu: row[2] || '', // Cột C
                    xuongSanXuat: row[3] || '', // Cột D
                    ngayTao: row[4] || '', // Cột E
                    thang: row[7] || '', // Cột H
                    nam: row[6] || '', // Cột G
                    nguoiTao: uriData.nguoiTaoURI || row[5] || '', // Cột F
                    ghiChu: row[11] || '', // Cột L
                };

                // Xác định tiêu đề dựa trên loại phiếu
                let title = "PHIẾU ĐỀ XÊ";
                if (orderDetails.loaiPhieu === "Xuất") {
                    title = "PHIẾU XUẤT KHO ĐỀ XÊ";
                } else if (orderDetails.loaiPhieu === "Nhập") {
                    title = "PHIẾU NHẬP KHO ĐỀ XÊ";
                }

                // Cập nhật nội dung HTML
                document.querySelector("h2").textContent = title;
                document.getElementById('maPhieu').textContent = orderDetails.maPhieu;
                document.getElementById('xuongSanXuat').textContent = orderDetails.xuongSanXuat;
                document.getElementById('ngayTao').textContent = extractDay(orderDetails.ngayTao);
                document.getElementById('thang').textContent = orderDetails.thang;
                document.getElementById('nam').textContent = orderDetails.nam;
                document.getElementById('nguoiTao').textContent = orderDetails.nguoiTao;
                document.getElementById('ghiChu').textContent = orderDetails.ghiChu;

                return; // Dừng khi tìm thấy
            }
        }

        updateContent(`No matching data found for "${maPhieuURI}".`);
    } catch (error) {
        updateContent('Error fetching data: ' + error.message);
        console.error('Fetch Error:', error);
    }

    function updateContent(message) {
        // Hàm để xử lý thông báo lỗi hoặc cập nhật chung
        alert(message);
    }
}

// Tìm chi tiết trong bảng tính
async function findDetailsInSheet(maPhieuURI) {
    try {
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: RANGE_CHITIET,
        });

        const rows = response.result.values;
        if (!rows || rows.length === 0) {
            updateContent('No detail data found.');
            return;
        }

        const filteredRows = rows.filter(row => row[0] === maPhieuURI); // Lọc các dòng có giá trị cột F khớp với maPhieuURI
        orderItems = filteredRows.map(extractDetailDataFromRow);
        if (filteredRows.length > 0) {
            displayDetailData(filteredRows);
        } else {
            updateContent('No matching data found.');
        }
    } catch (error) {
        console.error('Error fetching detail data:', error);
        updateContent('Error fetching detail data.');
    }
}

function displayDetailData(filteredRows) {
    const tableBody = document.getElementById('itemTableBody');
    tableBody.innerHTML = ''; // Xóa dữ liệu cũ nếu có

    filteredRows.forEach(row => {
        const item = extractDetailDataFromRow(row);;

        tableBody.innerHTML += `
        <tr class="bordered-table">
            <td class="borderedcol-1">${item.nhomVattu || ''}</td>
            <td class="borderedcol-2">${item.maVattu || ''}</td>
            <td class="borderedcol-3">${item.tenVattu || ''}</td>
            <td class="borderedcol-4">${item.chieuDai || ''}</td>
            <td class="borderedcol-5">${item.soNep || ''}</td>
            <td class="borderedcol-6">${item.slXuatQuydoi || ''}</td>
            <td class="borderedcol-7">${item.dvtQuydoi || ''}</td>
            <td class="borderedcol-8">${item.chieuDainhaplai || ''}</td>
            <td class="borderedcol-9">${item.soNepnhaplai || ''}</td>
            <td class="borderedcol-10">${item.vitriKehang || ''}</td>
            <td class="borderedcol-11">${item.ghiChuItem || ''}</td>
        </tr>
    `;
    });
}


// Trích xuất dữ liệu từ hàng
function extractDetailDataFromRow(row) {
    return {
        nhomVattu: row[8],
        maVattu: row[9],
        tenVattu: row[10],
        chieuDai: row[11],
        soNep: row[12],
        slXuatQuydoi: row[15],
        dvtQuydoi: row[16],
        chieuDainhaplai: row[17],
        soNepnhaplai: row[18],
        vitriKehang: row[19],
        ghiChuItem: row[20],
    };
}

// Hàm cập nhật nội dung DOM
function updateElement(elementId, value) {
    const element = document.getElementById(elementId);
    if (element) {
        element.innerText = value;
    }
}