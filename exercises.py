import heapq
import math
from openpyxl import load_workbook

# =====================================================
#  ระบบแนะนำการออกกำลังกาย (ใช้ A* Search + Heuristic)
# =====================================================
# f(n) = g(n) + h(n)
# g(n) = "ความเหนื่อยสะสม" ที่เสียไปแล้วตั้งแต่เริ่มวางแผน
# h(n) = ค่าประมาณความเหนื่อยที่ "ยังต้องเสียอีก" เพื่อให้ได้แคลครบตามเป้า
# f(n) = ความเหนื่อยรวมโดยคาดการ  (A* จะเลือกแผนที่ f น้อยสุดก่อนเสมอ)
#
# state (n) = (วันที่, ท่าที่เล่นไปแล้ววันนี้, แคลที่เผาผลาญสะสม)
# goal      = เผาผลาญแคลได้ครบตามเป้าหมาย

FILE_NAME = "exercises.xlsx"

DAY_NAMES = ["จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์", "เสาร์", "อาทิตย์"]

# ตำแหน่งวันออกกำลังกายในสัปดาห์ (0 = จันทร์) จัดให้มีวันพักคั่นเสมอ
WEEK_PATTERN = {
    1: [0],
    2: [0, 3],
    3: [0, 2, 4],
    4: [0, 1, 3, 5],
    5: [0, 1, 3, 4, 5],
}

# โรคที่ให้เลือก (ชื่อโรค, คำที่ไปค้นในช่อง "เงื่อนไข/ข้อควรระวัง" ของ excel)
DISEASES = {
    1: ("ความดันโลหิตสูง", ["ความดันโลหิตสูง"]),
    2: ("โรคหัวใจ", ["โรคหัวใจ"]),
    3: ("เบาหวาน", ["เบาหวาน"]),
    4: ("ปวดเข่า / ข้อเข่าเสื่อม", ["ปวดเข่า", "ข้อเข่าเสื่อม"]),
    5: ("ปวดหลัง / หมอนรองกระดูกทับเส้น", ["ปวดหลัง", "หมอนรองกระดูก"]),
    6: ("ปวดคอ", ["ปวดคอ"]),
    7: ("เจ็บไหล่", ["ไหล่"]),
    8: ("ปัญหาข้อเท้า / กระดูกเปราะ", ["ข้อเท้า", "กระดูกเปราะ"]),
    9: ("เจ็บข้อมือ", ["ข้อมือ"]),
}

# ข้อความที่แปลว่า "ท่านี้ปลอดภัย" ไม่ได้ห้ามใคร
SAFE_WORDS = ["เหมาะ", "ปลอดภัย", "แรงกระแทกต่ำ", "ใช้ดัมเบล"]
# ข้อความที่แปลว่า "เล่นได้ แต่ให้ลดลง" (ไม่ได้ห้าม)
SOFT_WORDS = ["ให้ลด", "ให้น้อยลง", "ไม่ควรกลั้นหายใจ"]

DAY_COST = 5  # ค่าความเหนื่อย/ความยุ่งยาก ของการเพิ่มวันออกกำลังกาย 1 วัน


# -----------------------------------------------------
# 1) อ่านข้อมูลท่าออกกำลังกายจากไฟล์ excel
# -----------------------------------------------------
def read_exercises(filename):
    wb = load_workbook(filename)
    ws = wb["exercises"]

    data = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        age_min, age_max = str(row[0]).split("-")
        kcal = float(str(row[5]).replace("~", "").replace("kcal", "").strip())
        data.append({
            "age_min": int(age_min),
            "age_max": int(age_max),
            "name": row[1],
            "muscle": row[2],
            "sets": int(row[3]),
            "reps": row[4],
            "kcal_set": kcal,
            "note": str(row[6]),
        })
    return data


# -----------------------------------------------------
# 2) คัดท่าที่เหมาะกับอายุและโรคประจำตัว
# -----------------------------------------------------
def filter_exercises(all_ex, age, my_diseases):
    age_used = min(age, 59)  # ไฟล์มีข้อมูลถึง 59 ปี ถ้าอายุมากกว่านี้ใช้ชุดเดียวกับ 40-59

    ok_list = []
    banned = []

    for ex in all_ex:
        if not (ex["age_min"] <= age_used <= ex["age_max"]):
            continue

        ex = dict(ex)  # copy จะได้ไม่แก้ข้อมูลต้นฉบับ
        ex["warning"] = ""
        note = ex["note"]

        is_safe = any(note.startswith(w) for w in SAFE_WORDS)
        is_soft = any(w in note for w in SOFT_WORDS)

        is_banned = False
        if not is_safe:
            for d in my_diseases:
                disease_name, keywords = DISEASES[d]
                if any(k in note for k in keywords):
                    if is_soft:
                        # เล่นได้ แต่ลดจำนวนเซ็ตลง 1
                        ex["sets"] = max(1, ex["sets"] - 1)
                        ex["warning"] = f"มี{disease_name} -> ลดเหลือ {ex['sets']} เซ็ต ({note})"
                    else:
                        is_banned = True
                        banned.append((ex["name"], disease_name))
                    break

        if is_banned:
            continue

        # เตรียมค่าที่ A* จะใช้
        k = ex["kcal_set"]
        ex["cal"] = ex["sets"] * k                    # แคลรวมของท่านี้ (ทุกเซ็ต)
        ex["effort"] = ex["sets"] * (1 + (k / 5) ** 2)  # ความเหนื่อย (ท่ายิ่งหนัก ยิ่งเหนื่อยเร็วขึ้น)
        ok_list.append(ex)

    return ok_list, banned


# -----------------------------------------------------
# 3) Heuristic: กำหนดจำนวนวันสูงสุด และจำนวนเซ็ตสูงสุดต่อวัน
# -----------------------------------------------------
def find_limit(age, bmi, my_diseases):
    if age <= 13:
        max_days, max_sets = 3, 8
    elif age <= 19:
        max_days, max_sets = 4, 12
    elif age <= 39:
        max_days, max_sets = 5, 16
    elif age <= 59:
        max_days, max_sets = 4, 10
    else:
        max_days, max_sets = 3, 8

    for d in my_diseases:
        if d == 2:              # โรคหัวใจ ลดมากกว่าปกติ
            max_days -= 2
        else:
            max_days -= 1
    if my_diseases:
        max_sets = max_sets * 0.8   # มีโรค ลดปริมาณต่อวัน 20%

    if bmi >= 30:               # น้ำหนักเยอะ ข้อต่อรับภาระมาก
        max_days -= 1

    max_days = max(max_days, 2)  # อย่างน้อย 2 วัน/สัปดาห์
    max_sets = max(int(max_sets), 4)
    return max_days, max_sets


# -----------------------------------------------------
# 4) A* Search
# -----------------------------------------------------
def a_star(exercises, target, max_days, max_sets):
    n = len(exercises)

    # ค่า "ความเหนื่อยต่อ 1 แคล" ที่ต่ำที่สุด ใช้ทำ heuristic
    # (ไม่ว่าเลือกท่าไหน ก็เหนื่อยไม่น้อยกว่านี้ -> h ไม่เกินความจริง)
    best_ratio = min(ex["effort"] / ex["cal"] for ex in exercises)

    def h(cal_now):
        return max(0.0, target - cal_now) * best_ratio

    # (f, g, วัน, ท่าที่เล่นวันนี้(bitmask), แคลสะสม, แผน)
    queue = [(h(0.0), 0.0, 0, 0, 0.0, ((),))]
    visited = {}

    while queue:
        f, g, day, mask, cal, plan = heapq.heappop(queue)

        if cal >= target:
            return plan, g, cal

        state = (day, mask, round(cal, 1))
        if state in visited and visited[state] <= g:
            continue
        visited[state] = g

        # จำนวนเซ็ตที่เล่นไปแล้วในวันนี้
        sets_today = 0
        for i in range(n):
            if mask & (1 << i):
                sets_today += exercises[i]["sets"]

        # ทางเลือก 1: เพิ่มท่าใหม่ในวันนี้
        for i in range(n):
            if mask & (1 << i):
                continue  # ท่านี้เล่นไปแล้ววันนี้
            if sets_today + exercises[i]["sets"] > max_sets:
                continue  # เกินโควต้าต่อวัน (กันออกกำลังกายหนักเกินไป)

            fatigue = 1 + 0.05 * sets_today  # เล่นมาเยอะแล้ว ท่าต่อไปยิ่งเหนื่อย
            new_g = g + exercises[i]["effort"] * fatigue
            new_cal = cal + exercises[i]["cal"]
            new_plan = plan[:-1] + (plan[-1] + (i,),)
            new_f = new_g + h(new_cal)
            heapq.heappush(queue, (new_f, new_g, day, mask | (1 << i), new_cal, new_plan))

        # ทางเลือก 2: จบวันนี้ ไปออกกำลังกายวันถัดไป
        if mask != 0 and day + 1 < max_days:
            new_g = g + DAY_COST
            new_plan = plan + ((),)
            new_f = new_g + h(cal)
            heapq.heappush(queue, (new_f, new_g, day + 1, 0, cal, new_plan))

    return None, 0, 0


# -----------------------------------------------------
# 5) ฟังก์ชันรับ input
# -----------------------------------------------------
def ask_number(text, min_value, max_value):
    while True:
        try:
            value = float(input(text))
            if min_value <= value <= max_value:
                return value
            print(f"  กรุณากรอกค่าระหว่าง {min_value} - {max_value}")
        except ValueError:
            print("  กรุณากรอกเป็นตัวเลข")


def ask_diseases():
    print("\nโรคประจำตัว (เลือกได้หลายข้อ พิมพ์เลขคั่นด้วย , เช่น 1,4)")
    print("  0. ไม่มีโรคประจำตัว")
    for num, (name, _) in DISEASES.items():
        print(f"  {num}. {name}")

    while True:
        text = input("เลือก: ").replace(" ", "")
        if text == "0":
            return []
        try:
            chosen = [int(x) for x in text.split(",")]
            if all(c in DISEASES for c in chosen):
                return sorted(set(chosen))
        except ValueError:
            pass
        print("  กรุณาพิมพ์เลขตามเมนู")


# -----------------------------------------------------
# 6) main
# -----------------------------------------------------
print("===== ระบบแนะนำการออกกำลังกาย =====")
age = int(ask_number("อายุ (ปี): ", 7, 100))
weight = ask_number("น้ำหนัก (กก.): ", 10, 300)
height = ask_number("ส่วนสูง (ซม.): ", 80, 250)
my_diseases = ask_diseases()
target_cal = ask_number("\nปริมาณแคลที่อยากลด (kcal): ", 1, 100000)

# คำนวณ BMI
bmi = weight / ((height / 100) ** 2)
print(f"\nBMI = {bmi:.1f}", end=" ")
if bmi < 18.5:
    print("(น้ำหนักต่ำกว่าเกณฑ์ ไม่แนะนำให้ลดน้ำหนักเพิ่ม ควรปรึกษาแพทย์)")
elif bmi < 23:
    print("(สมส่วน)")
elif bmi < 25:
    print("(น้ำหนักเกินเล็กน้อย)")
elif bmi < 30:
    print("(อ้วนระดับ 1)")
else:
    print("(อ้วนระดับ 2 ระวังเรื่องข้อเข่า/ข้อเท้า)")

if age > 59:
    print("หมายเหตุ: ไฟล์มีข้อมูลถึงอายุ 59 ปี จึงใช้ท่าของกลุ่ม 40-59 และลดจำนวนวันลง")

# อ่าน excel + คัดท่า
all_ex = read_exercises(FILE_NAME)
exercises, banned = filter_exercises(all_ex, age, my_diseases)

if banned:
    print("\nท่าที่ไม่แนะนำสำหรับคุณ:")
    for name, why in banned:
        print(f"  - {name} (เพราะ {why})")

if len(exercises) == 0:
    print("\nไม่มีท่าที่เหมาะสม กรุณาปรึกษาแพทย์ก่อนออกกำลังกาย")
    raise SystemExit

# หาจำนวนวัน/โควต้าต่อวัน
max_days, max_sets = find_limit(age, bmi, my_diseases)

# หาว่า 1 วัน เผาผลาญได้มากสุดเท่าไหร่ (ภายใต้โควต้าเซ็ต) -> ดูว่า 1 สัปดาห์ทำได้กี่แคล
best_day_cal = 0
for mask in range(1 << len(exercises)):
    s = 0
    c = 0
    for i in range(len(exercises)):
        if mask & (1 << i):
            s += exercises[i]["sets"]
            c += exercises[i]["cal"]
    if s <= max_sets and c > best_day_cal:
        best_day_cal = c

week_capacity = best_day_cal * max_days

# ถ้าเป้าหมายมากกว่าที่ปลอดภัยต่อสัปดาห์ ให้แบ่งเป็นหลายสัปดาห์
weeks = math.ceil(target_cal / week_capacity)
week_target = target_cal / weeks

route, total_effort, total_cal = a_star(exercises, week_target, max_days, max_sets)

if route is None:
    print("ไม่พบแผนที่เหมาะสม")
    raise SystemExit

# จัดลงตารางสัปดาห์ (เอาวันที่ว่างเปล่าออก)
route = [d for d in route if len(d) > 0]
n_days = len(route)
positions = WEEK_PATTERN[n_days]

print("\n===== แผนออกกำลังกาย 1 สัปดาห์ =====")
print(f"ออกกำลังกาย {n_days} วัน / พัก {7 - n_days} วัน")

for day_index in range(7):
    if day_index in positions:
        todo = route[positions.index(day_index)]
        day_cal = sum(exercises[i]["cal"] for i in todo)
        print(f"\nวัน{DAY_NAMES[day_index]}  (เผาผลาญประมาณ {day_cal:.1f} kcal)")
        for i in todo:
            ex = exercises[i]
            print(f"  - {ex['name']} ({ex['muscle']}): {ex['sets']} เซ็ต x {ex['reps']}"
                  f"  ~{ex['cal']:.1f} kcal")
            if ex["warning"]:
                print(f"      ** {ex['warning']}")
    else:
        print(f"\nวัน{DAY_NAMES[day_index]}  พักผ่อน")

print("\n===== สรุป =====")
print(f"เผาผลาญได้สัปดาห์ละ ~{total_cal:.1f} kcal")
if weeks > 1:
    print(f"เป้าหมาย {target_cal:.0f} kcal มากเกินกว่าจะทำในสัปดาห์เดียวอย่างปลอดภัย")
    print(f"แนะนำให้ทำตามแผนนี้ซ้ำ {weeks} สัปดาห์")
else:
    print(f"ครบเป้าหมาย {target_cal:.0f} kcal ภายใน 1 สัปดาห์")
print("(ค่าความเหนื่อยรวมของแผนตาม A* = {:.2f})".format(total_effort))
print("\n*** ข้อมูลนี้เป็นเพียงคำแนะนำเบื้องต้น หากมีโรคประจำตัวควรปรึกษาแพทย์ ***")
