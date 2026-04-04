const FORM_RESPONSE_URL =
    "https://docs.google.com/forms/d/e/1FAIpQLSf2FkCXDl8bp_vVaHx2MvAstmXB-laXUO5s3F-2EZpb15sKHA/formResponse";

const SUBMIT_DELAY_MS = 1000;

const OPTIONS = {
    gender: ["Nam", "Nữ"],
    experience: [
        "Đoàn sinh",
        "Đã từng làm TNV và Thiền sinh",
        "Đã từng làm TNV",
        "Đã từng làm Thiền sinh",
        "Chưa từng làm TNV và Thiền sinh",
    ],
    transport: ["Đăng ký đi ô tô cùng Đoàn", "Tự túc phương tiện"],
    role: [
        "1. Hướng dẫn: Hướng dẫn thiền sinh khóa tu theo sự chỉ dạy của quý Thầy Cô, điều phối từ Quản chúng và Ban tổ chức.",
        "2. Hậu cần: Sơ chế/Hành đường/Bếp",
        "3. Truyền thông: Chụp hình hoặc quay dựng video (yêu cầu mang theo máy ảnh cá nhân)",
    ],
};

function submitForm100Times() {
    for (let i = 1; i <= 100; i++) {
        const payload = buildPayload(i);

        UrlFetchApp.fetch(FORM_RESPONSE_URL, {
            method: "post",
            payload,
            followRedirects: true,
            muteHttpExceptions: true,
        });

        if (i < 100) Utilities.sleep(SUBMIT_DELAY_MS);
    }
}

function buildPayload(i) {
    const role = pick(OPTIONS.role);

    return {
        "entry.1816309196": `Auto TNV ${String(i).padStart(3, "0")}`, // Họ tên
        "entry.1520738930": pick(OPTIONS.gender), // Giới tính
        "entry.473785259": `auto.tnv.${i}@example.com`, // Email (new field, optional)
        "entry.1967478537": randomDob(), // Ngày sinh dd/MM/yyyy
        "entry.1238559497": randomPhone(), // SĐT
        "entry.1329755398": pick(OPTIONS.experience), // Kinh nghiệm
        "entry.1384622437": pick(OPTIONS.transport), // Di chuyển
        "entry.1515842689": "Đồng ý", // Đồng ý gửi điện thoại
        "entry.445178271": role, // Vị trí mong muốn
        "entry.1350537498": `Mình phù hợp với vai trò: ${role}`, // Lý do phù hợp
        "entry.1536782630": "", // Câu hỏi thêm (optional)
    };
}

function pick(arr) {
    return arr[Math.floor(Math.random() * arr.length)];
}

function randomDob() {
    const year = randInt(1990, 2005);
    const month = randInt(1, 12);
    const day = randInt(1, 28); // safe for every month
    return `${String(day).padStart(2, "0")}/${String(month).padStart(2, "0")}/${year}`;
}

function randInt(min, max) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
}

function randomPhone() {
    return "09" + Math.floor(10000000 + Math.random() * 90000000);
}