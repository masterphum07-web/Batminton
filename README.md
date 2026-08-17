# Freshy Badminton Scoreboard

## เชื่อม Google Sheets

1. เปิด Google Sheet แล้วไปที่ **ส่วนขยาย > Apps Script**
2. วางโค้ดจาก `Code.gs` แทนโค้ดเดิม แก้ `API_KEY` เป็นรหัสลับของคุณ แล้วกด `setupSheet()` หนึ่งครั้งเพื่อสร้างชีต `Matches`
3. กด **Deploy > New deployment > Web app** เลือก Execute as: Me และ Who has access: Anyone แล้วกด Deploy
4. คัดลอก URL ที่ลงท้ายด้วย `/exec`
5. เปิด `index.html` แล้วใส่ URL ใน `SHEETS_API_URL` และใส่ API key เดียวกันใน `SHEETS_API_KEY`

## ทดสอบเว็บ

เปิด `index.html` ในเบราว์เซอร์ หากต้องการเผยแพร่ ให้สร้าง repository บน GitHub แล้วอัปโหลด `index.html`, `Code.gs`, `README.md` จากนั้นเปิด GitHub Pages จาก Settings > Pages > Deploy from branch > main/root

หมายเหตุ: ข้อมูลคะแนนจะถูกเก็บในชีต `Matches` โดยแต่ละแมตช์เป็นหนึ่งแถว และมี version ป้องกันการบันทึกทับกัน
