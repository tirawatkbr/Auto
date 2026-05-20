# RevPi Auto-Fill — คู่มือติดตั้งและใช้งาน

ระบบกรอกข้อมูลอัตโนมัติ: ดึงไฟล์ ST99 Production Report จากเมล Comanche
แล้วกรอกค่าลงไฟล์ RevPiData (.xlsx) ในเครื่อง โดยรักษาสูตรและรูปแบบไว้ครบ

---

## ภาพรวมการทำงาน

```
   เมล Comanche          ระบบนี้                ไฟล์ในเครื่อง
  ┌────────────┐      ┌──────────────┐       ┌──────────────┐
  │ ST99       │ ───> │ ดาวน์โหลด    │ ───>  │ RevPiData    │
  │ .xlsx แนบ  │      │ parse + กรอก │       │ .xlsx        │
  └────────────┘      │ ตรวจยอดรวม   │       │ (กรอกแล้ว)   │
                      └──────────────┘       └──────────────┘
                            │
                            ├─> สำรองไฟล์เดิมไว้ใน backups/
                            └─> บันทึก log ใน logs/
```

ทุกครั้งที่รัน ระบบจะ:
1. ค้นเมลจาก Comanche ที่มีไฟล์แนบ ST99 (เฉพาะเมลใหม่ที่ยังไม่เคยทำ)
2. ดาวน์โหลดไฟล์แนบ
3. เดาเดือนปลายทางจากช่วงวันที่ในรายงาน
4. **สำรองไฟล์ RevPiData เดิมก่อนเสมอ**
5. **ตรวจสอบยอดรวมก่อนเขียน** — ถ้ายอดไม่ตรง Grand Total จะไม่เขียนไฟล์
6. กรอกข้อมูลลงบล็อกเดือนที่ตรง (ช่อง RNs / REV / F&B)
7. บันทึก log

---

## ขั้นตอนที่ 1 — ติดตั้ง Python และไลบรารี

### 1.1 ติดตั้ง Python
ถ้ายังไม่มี Python ดาวน์โหลดจาก https://www.python.org/downloads/
(เวอร์ชัน 3.9 ขึ้นไป — ตอนติดตั้งบน Windows **ติ๊กถูก "Add Python to PATH"**)

### 1.2 ติดตั้งไลบรารีที่ต้องใช้
เปิด Command Prompt (Windows) หรือ Terminal (Mac) แล้วพิมพ์:

```
cd path/ไปยัง/โฟลเดอร์ revpi_autofill
pip install -r requirements.txt
```

---

## ขั้นตอนที่ 2 — สร้าง Gmail API Credentials (ทำครั้งเดียว)

ระบบต้องขออนุญาตอ่านเมลของคุณ ผ่าน Google Cloud (ฟรี)

### 2.1 สร้างโปรเจกต์ใน Google Cloud
1. เข้า https://console.cloud.google.com/
2. กดเมนูบนซ้าย > **เลือกโปรเจกต์** > **โปรเจกต์ใหม่ (New Project)**
3. ตั้งชื่อ เช่น `RevPi AutoFill` > **สร้าง**

### 2.2 เปิดใช้ Gmail API
1. ที่แถบค้นหาด้านบน พิมพ์ **Gmail API** > กดเข้าไป
2. กดปุ่ม **เปิดใช้ (Enable)**

### 2.3 ตั้งค่า OAuth Consent Screen
1. เมนูซ้าย > **APIs & Services** > **OAuth consent screen**
2. เลือก **External** > **สร้าง**
3. กรอกชื่อแอป (เช่น `RevPi AutoFill`), อีเมลของคุณในช่องที่บังคับ > **บันทึกและไปต่อ**
4. หน้า Scopes — กด **บันทึกและไปต่อ** (ข้ามได้)
5. หน้า Test users — กด **+ ADD USERS** ใส่อีเมล Gmail ของคุณเอง > **บันทึก**

### 2.4 สร้าง Credentials
1. เมนูซ้าย > **APIs & Services** > **Credentials**
2. กด **+ CREATE CREDENTIALS** > **OAuth client ID**
3. Application type เลือก **Desktop app** > ตั้งชื่อ > **สร้าง**
4. กดปุ่ม **ดาวน์โหลด JSON**
5. **เปลี่ยนชื่อไฟล์ที่ดาวน์โหลดเป็น `credentials.json`**
   แล้ววางไว้ในโฟลเดอร์ `revpi_autofill` (โฟลเดอร์เดียวกับ run.py)

> หมายเหตุ: ระบบขอสิทธิ์แบบ **อ่านอย่างเดียว (readonly)** — ลบหรือส่งเมลไม่ได้

---

## ขั้นตอนที่ 3 — ตั้งค่าระบบในไฟล์ config.py

เปิดไฟล์ `config.py` ด้วย Notepad / VS Code แล้วแก้ 2 จุดหลัก:

### 3.1 ระบุไฟล์ปลายทาง
หาบรรทัด `TARGET_XLSX` แล้วใส่ path เต็มของไฟล์ RevPiData ในเครื่อง:

```python
# Windows
TARGET_XLSX = r"C:\Users\YIM\Documents\RevPiData-R18.xlsx"

# Mac
TARGET_XLSX = "/Users/yim/Documents/RevPiData-R18.xlsx"
```

### 3.2 ตั้งเงื่อนไขกรองเมล
หาส่วน "การกรองเมลจาก Comanche" แล้วปรับ:

```python
GMAIL_QUERY_FROM      = "noreply@comanche.com"  # อีเมลผู้ส่งของ Comanche
GMAIL_SUBJECT_KEYWORD = "ST99"                  # คำในหัวข้อเมล
GMAIL_NEWER_THAN      = "7d"                    # ดูเมลย้อนหลังกี่วัน
ATTACHMENT_NAME_PREFIX = "ST99"                 # ชื่อไฟล์แนบขึ้นต้นด้วย
```

> **สำคัญ:** ถ้าไม่แน่ใจอีเมลผู้ส่ง ให้เปิด Gmail ดูเมลจริงจาก Comanche
> ก่อน แล้วคัดลอกอีเมลผู้ส่งมาใส่ — หรือเว้น `GMAIL_QUERY_FROM = ""`
> ไว้ก่อน แล้วใช้แค่ keyword หัวข้อกรอง

---

## ขั้นตอนที่ 4 — ทดสอบก่อนใช้จริง

**สำคัญ:** ก่อนใช้กับไฟล์จริง ให้ทดสอบตามคู่มือ `TESTING_GUIDE.md`
ซึ่งพาทดสอบเป็น 5 ด่าน เรียงจากปลอดภัยสุดไปหาเต็มรูปแบบ

เริ่มจากตรวจความพร้อมระบบก่อน:

```
python doctor.py
```

ระบบจะตรวจทุกอย่าง (Python, ไลบรารี, ไฟล์ปลายทาง, credentials,
เงื่อนไขกรองเมล) แล้วบอกว่าอะไรพร้อม/ไม่พร้อม โดยไม่แตะไฟล์ใดๆ

จากนั้นทำตาม `TESTING_GUIDE.md` ด่าน 2–5 ต่อ

---

## ขั้นตอนที่ 5 — ตั้งให้รันอัตโนมัติเต็มรูปแบบ

### บน Windows — ใช้ Task Scheduler
1. เปิด **Task Scheduler** (ค้นใน Start menu)
2. กด **Create Basic Task** ทางขวา
3. ตั้งชื่อ เช่น `RevPi AutoFill` > **Next**
4. เลือกความถี่ เช่น **Daily** > ตั้งเวลา (เช่น 08:00) > **Next**
5. เลือก **Start a program** > **Next**
6. ช่อง Program/script กด **Browse** เลือกไฟล์ `run_windows.bat`
7. **Next** > **Finish**

> แนะนำ: ดับเบิลคลิก task ที่สร้าง > แท็บ General >
> ติ๊ก **Run whether user is logged on or not**

### บน Mac — ใช้ cron
1. เปิด Terminal พิมพ์ `crontab -e`
2. เพิ่มบรรทัด (รันทุกวัน 08:00 — แก้ path ให้ตรงจริง):

```
0 8 * * * /bin/bash /Users/yim/revpi_autofill/run_mac.sh
```

3. บันทึกและออก (กด `Esc` แล้วพิมพ์ `:wq` ถ้าใช้ vi)

> ก่อนใช้ ให้สั่ง `chmod +x run_mac.sh` หนึ่งครั้ง

---

## โครงสร้างไฟล์ในโปรเจกต์

```
revpi_autofill/
├── config.py            ← ตั้งค่าทั้งหมด (แก้ไฟล์นี้)
├── run.py               ← ตัวหลัก
├── filler.py            ← parse ST99 + กรอก Excel
├── gmail_client.py      ← เชื่อม Gmail
├── doctor.py            ← ตรวจความพร้อมระบบ
├── requirements.txt     ← รายการไลบรารี
├── run_windows.bat      ← ตัวรันสำหรับ Windows
├── run_mac.sh           ← ตัวรันสำหรับ Mac
├── SETUP_GUIDE.md       ← คู่มือติดตั้ง (ไฟล์นี้)
├── TESTING_GUIDE.md     ← คู่มือทดสอบก่อนใช้จริง
├── credentials.json     ← (คุณวางเอง จาก Google Cloud)
├── token.json           ← (สร้างอัตโนมัติรอบแรก)
├── processed.json       ← (สร้างอัตโนมัติ — บันทึกเมลที่ทำแล้ว)
├── backups/             ← สำรองไฟล์ RevPiData ทุกครั้งก่อนเขียน
├── logs/                ← บันทึกการทำงานรายวัน
└── _inbox/              ← ไฟล์แนบที่ดาวน์โหลดมา
```

---

## กลไกความปลอดภัย (ออกแบบมาให้รันอัตโนมัติได้อย่างมั่นใจ)

| กลไก | ทำอะไร |
|------|--------|
| สำรองไฟล์ | คัดลอก RevPiData เดิมไป backups/ ก่อนเขียนทุกครั้ง |
| ตรวจยอดก่อนเขียน | ถ้ายอดรวมต่างจาก Grand Total เกินเกณฑ์ จะไม่เขียนไฟล์เลย |
| ป้องกันเซลล์สูตร | ไม่เขียนทับเซลล์ที่เป็นสูตร เด็ดขาด |
| กันกรอกซ้ำ | จดจำเมลที่ทำแล้วใน processed.json |
| Log ครบ | บันทึกทุกการทำงานใน logs/ ย้อนดูได้ |

ถ้าผลออกมาผิด — เปิดไฟล์ใน `backups/` มาทับคืนได้ทันที

---

## การปรับแต่งที่พบบ่อย

**บังคับเดือนปลายทาง** (กรณีระบบเดาผิด):
```
python run.py --month MAY
```

**ถ้ายอดรวมไม่ตรงแล้วอยากให้เขียนต่อ** (ไม่แนะนำ):
แก้ใน config.py: `ABORT_ON_VERIFY_FAIL = False`

**ถ้า ST99 เดือนอื่นมี segment ใหม่ที่ไม่อยู่ในตาราง:**
เพิ่มลงใน `SEGMENT_MAP` ใน config.py เช่น
```python
"NEW SEGMENT NAME": "ชื่อ leaf ปลายทาง",
```

**เปลี่ยนไปกรอกคอลัมน์ MONTH END แทน ON BOOKS:**
แก้ใน config.py: `ACTIVE_COLSET = "MONTHEND"`

---

## ตาราง Mapping ที่ตั้งไว้

| ST99 (Comanche) | RevPiData (leaf segment) |
|---|---|
| CORPORATE FIT | COF Corporate FIT |
| CORPORATE GROUP | COG Corporate Group |
| TRAVEL AGENT FIT | WHF Wholesale FIT |
| TRAVEL AGENT GROUP | WHG Wholesale Group |
| ONLINE FIT OTA | OTA |
| ONLINE HOTEL WEB FIT | HOTEL WEB |
| DIRECT HOTEL | DIR Direct / Walk In |
| STAFF RATE | COO Owner |
| OTHERS COMPLIMENTARY | Complimentary |
| AIRLINE CREW | ALC Airline Crew |

---

## แก้ปัญหาเบื้องต้น

| อาการ | วิธีแก้ |
|------|--------|
| `ไม่พบไฟล์ credentials.json` | ทำขั้นตอนที่ 2 ให้ครบ วางไฟล์ในโฟลเดอร์ที่ถูก |
| `ไม่พบไฟล์ปลายทาง` | แก้ `TARGET_XLSX` ใน config.py ให้เป็น path จริง |
| `ไม่พบเมลใหม่` | ตรวจ `GMAIL_QUERY_FROM` / keyword / `GMAIL_NEWER_THAN` |
| `เดาเดือนไม่ได้` | รันด้วย `--month MAY` ระบุเดือนเอง |
| `ยอดรวมไม่ผ่านเกณฑ์` | ดู log ว่ามี segment จับคู่ไม่ได้ไหม — เพิ่มใน SEGMENT_MAP |
| token หมดอายุ | ลบ `token.json` แล้วรันใหม่ ล็อกอินอีกครั้ง |

ดู log ละเอียดได้ในโฟลเดอร์ `logs/`
