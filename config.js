// ตั้งค่าการเชื่อมต่อ: วาง URL ของ Apps Script Web app (ลงท้ายด้วย /exec) ที่ได้จากการ Deploy
window.SW_CONFIG = {
  API_URL: "https://script.google.com/macros/s/AKfycbzN767gHVqxl62af1IjV2VUhXHeJkQpOYtOURIi98W2bpby50p0TPD-r7Nibn35uFvQkQ/exec",          // เช่น "https://script.google.com/macros/s/XXXXXXXX/exec"
  POLL_SECONDS: 20      // ดึงข้อมูลใหม่ทุกกี่วินาที
};
