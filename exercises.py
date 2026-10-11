import heapq
import math
import os
import re
from itertools import combinations

import openpyxl

# =====================================================================
#  ระบบแนะนำการออกกำลังกายด้วย A* Search
#  รับ: อายุ, น้ำหนัก, ส่วนสูง, โรคประจำตัว, แคลอรีที่ต้องการลด
#  อ่านท่าออกกำลังกายและเงื่อนไขจากไฟล์ exercises.xlsx
# =====================================================================

EXCEL_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "exercises.xlsx")   # วางไว้โฟลเดอร์เดียวกับไฟล์นี้
SHEET_NAME = "exercises"
MAX_WEEKS = 26          # วางแผนได้สูงสุด 26 สัปดาห์
WEEK_LOAD = 0.85        # สัปดาห์ที่ไม่ใช่สัปดาห์สุดท้าย ตั้งเป้า 85% ของความจุสูงสุดต่อสัปดาห์ (ไม่หนักเต็มเพดาน)

# ---------------------------------------------------------------------
#  ตารางกฎ Heuristic (ปรับค่าได้ตามต้องการ)
# ---------------------------------------------------------------------
# ค่าเริ่มต้นตามช่วงอายุ
#   max_days      = ออกกำลังกายได้สูงสุดกี่วันต่อสัปดาห์
#   max_consec    = ออกกำลังกายติดกันได้สูงสุดกี่วัน (แล้วต้องพัก)
#   max_kcal      = แคลอรีสูงสุดที่ควรเผาผลาญต่อวัน
#   max_exercises = จำนวนท่าสูงสุดต่อวัน
AGE_RULES = {
    (7, 13):  {"max_days": 4, "max_consec": 2, "max_kcal": 45,  "max_exercises": 3},
    (14, 19): {"max_days": 5, "max_consec": 3, "max_kcal": 70,  "max_exercises": 3},
    (20, 39): {"max_days": 5, "max_consec": 3, "max_kcal": 110, "max_exercises": 3},
    (40, 59): {"max_days": 4, "max_consec": 2, "max_kcal": 60,  "max_exercises": 3},
}

# โรคประจำตัว -> (คำที่ใช้ค้นในช่อง "เงื่อนไข/ข้อควรระวัง" ของ Excel,
#                 จำนวนวันที่ต้องลดต่อสัปดาห์, ตัวคูณแคลอรีต่อวัน)
DISEASE_MENU = {
    "ความดันโลหิตสูง":              (["ความดันโลหิตสูง"],         1, 0.85),
    "โรคหัวใจ":                      (["โรคหัวใจ"],                 2, 0.60),
    "เบาหวาน":                       (["เบาหวาน"],                  0, 0.85),
    "ปวดเข่า/ข้อเข่าเสื่อม":          (["เข่า"],                     1, 0.85),
    "ปวดหลัง/หมอนรองกระดูกทับเส้น":   (["หลัง", "หมอนรองกระดูก"],   1, 0.85),
    "ปวดไหล่":                       (["ไหล่"],                     0, 0.90),
    "ปวดคอ":                         (["ปวดคอ"],                    0, 0.90),
    "ข้อเท้า/กระดูกเปราะ":            (["ข้อเท้า", "กระดูกเปราะ"],   1, 0.80),
    "ปวดข้อมือ":                     (["ข้อมือ"],                   0, 0.90),
}

# คำในช่องเงื่อนไขของ Excel ที่ช่วยแยกประเภทข้อความ
SAFE_WORDS = ["ปลอดภัย", "เหมาะสำหรับ", "แรงกระแทกต่ำ", "ใช้ดัมเบล"]   # ข้อความแนะนำ ไม่ใช่ข้อห้าม
MODIFY_WORDS = ["ระวัง", "ไม่ควรกลั้นหายใจ", "ให้ลด", "น้อยลง"]        # เล่นได้แต่ต้องปรับ (ลด 1 เซ็ต)

# น้ำหนักค่าใช้จ่ายของ A* (ยิ่งมากยิ่ง "เหนื่อย/เสี่ยง")
STRAIN_KCAL_W = 0.5     # ความเหนื่อยตามปริมาณแคลที่เผาในวันนั้น
OVERLAP_W = 0.5         # กล้ามเนื้อกลุ่มเดิมซ้ำกับเมื่อวาน (กล้ามเนื้อยังไม่ฟื้น)
CONSEC_W = 0.2          # ออกกำลังกายติดกันหลายวัน ความล้าสะสม


# =====================================================================
#  1) อ่านข้อมูลจาก Excel
# =====================================================================
def load_exercises(path=EXCEL_FILE):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[SHEET_NAME]
    data = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        lo, hi = (int(x) for x in str(row[0]).split("-"))
        data.append({
            "age_lo": lo,
            "age_hi": hi,
            "name": row[1],
            "muscle": set(str(row[2]).split("/")),
            "muscle_text": row[2],
            "sets": int(row[3]),
            "reps": row[4],
            "kcal_per_set": float(re.search(r"[\d.]+", str(row[5])).group()),
            "note": str(row[6] or ""),
        })
    return data


# =====================================================================
#  2) กรองท่าตามอายุ และโรคประจำตัว
# =====================================================================
def classify(note, keywords):
    # คืนค่า "ok" (เล่นได้), "modify" (เล่นได้แต่ลดปริมาณ), "avoid" (ไม่ควรเล่น)
    if any(w in note for w in SAFE_WORDS):
        return "ok"
    if not any(k in note for k in keywords):
        return "ok"
    if any(w in note for w in MODIFY_WORDS):
        return "modify"
    return "avoid"


def filter_exercises(exercises, age, diseases):
    allowed, removed = [], []
    for ex in exercises:
        if not (ex["age_lo"] <= age <= ex["age_hi"]):
            continue
        status, why = "ok", []
        for d in diseases:
            s = classify(ex["note"], DISEASE_MENU[d][0])
            if s == "avoid":
                status = "avoid"
                why.append(d)
            elif s == "modify" and status != "avoid":
                status = "modify"
                why.append(d)
        if status == "avoid":
            removed.append((ex["name"], why))
            continue
        sets = ex["sets"] - 1 if status == "modify" else ex["sets"]
        sets = max(1, sets)
        allowed.append({**ex, "status": status, "sets": sets,
                        "kcal": sets * ex["kcal_per_set"]})
    return allowed, removed


# =====================================================================
#  3) กำหนดขีดจำกัดตามหลัก Heuristic (อายุ + BMI + โรค)
# =====================================================================
def build_limits(age, bmi, diseases):
    rule = next(v for (lo, hi), v in AGE_RULES.items() if lo <= age <= hi)
    limits = dict(rule)
    notes = []

    # ลดวัน/ลดแคลตามโรค
    cut_days, factor = 0, 1.0
    for d in diseases:
        _, dd, ff = DISEASE_MENU[d]
        cut_days += dd
        factor *= ff
    if diseases:
        limits["max_consec"] = min(limits["max_consec"], 2)
        notes.append("มีโรคประจำตัว: ลดจำนวนวันและแคลต่อวัน ไม่ออกติดกันเกิน 2 วัน")
    if "โรคหัวใจ" in diseases:
        limits["max_consec"] = 1
        notes.append("โรคหัวใจ: ออกกำลังกายวันเว้นวัน")

    # BMI ใช้เกณฑ์ผู้ใหญ่ จึงใช้กับอายุ 20 ปีขึ้นไปเท่านั้น
    if age >= 20:
        if bmi >= 30:
            cut_days += 1
            factor *= 0.85
            notes.append("BMI >= 30: ลดภาระต่อข้อต่อ ลด 1 วัน และลดแคลต่อวัน")
        elif bmi < 18.5:
            notes.append("BMI < 18.5 (น้ำหนักต่ำกว่าเกณฑ์): ไม่แนะนำให้ลดน้ำหนักเพิ่ม ควรปรึกษาแพทย์")
    else:
        notes.append("อายุต่ำกว่า 20 ปี: ไม่ใช้เกณฑ์ BMI ผู้ใหญ่ในการปรับแผน")

    # อย่างน้อย 3 วัน/สัปดาห์ เพื่อให้ยังได้ผลและไม่ยืดเวลาเกินไป
    limits["max_days"] = max(3, limits["max_days"] - cut_days)
    limits["max_kcal"] = limits["max_kcal"] * max(0.5, factor)
    return limits, notes


# =====================================================================
#  4) สร้าง "ชุดออกกำลังกายต่อวัน" (session) จากท่าที่เล่นได้
# =====================================================================
def build_sessions(exs, limits):
    sessions = []
    n = len(exs)
    for size in range(min(2, n), min(limits["max_exercises"], n) + 1):
        for combo in combinations(exs, size):
            kcal = sum(e["kcal"] for e in combo)
            if kcal <= limits["max_kcal"]:
                muscles = set().union(*(e["muscle"] for e in combo))
                sessions.append({"exs": list(combo), "kcal": kcal, "muscle": muscles})
    if not sessions:   # ถ้าเกินเพดานทุกชุด ให้ใช้ท่าเดี่ยวที่เบาที่สุด
        e = min(exs, key=lambda x: x["kcal"])
        sessions.append({"exs": [e], "kcal": e["kcal"], "muscle": set(e["muscle"])})
    return sessions


# =====================================================================
#  5) A* Search
# =====================================================================
# state n = (day, kcal, consec, week_count, last)
#   day        = วันที่ผ่านไปแล้ว
#   kcal       = แคลที่เผาผลาญสะสมแล้ว
#   consec     = จำนวนวันที่ออกกำลังกายติดกันล่าสุด
#   week_count = จำนวนวันที่ออกกำลังกายในสัปดาห์ปัจจุบัน
#   last       = index ของชุดออกกำลังกายเมื่อวาน (-1 = พัก)
#
# f(n) = g(n) + h(n)
# g(n) = ต้นทุนจริงที่ใช้ไปแล้ว (1 ต่อวัน + ความเหนื่อย/ความล้า)
# h(n) = ค่าประมาณต้นทุนที่เหลือ (จำนวนวันขั้นต่ำ + ความเหนื่อยขั้นต่ำ)
# f(n) = ต้นทุนรวมที่คาดการณ์

def heuristic(n, target, max_session_kcal, limits):
    day, kcal, consec, week_count, last = n
    rem = target - kcal
    if rem <= 0:
        return 0.0

    # วันออกกำลังกายขั้นต่ำที่ต้องใช้ ถ้าทุกวันทำได้เต็มที่
    k = math.ceil(rem / max_session_kcal)

    # จำกัดโควตาวันออกกำลังกายต่อสัปดาห์ (มองแบบมองโลกในแง่ดีที่สุด จึงไม่เกินจริง)
    w = limits["max_days"]
    days_left = 7 - day % 7
    q = min(w - week_count, days_left)
    if k <= q:
        min_days = k
    else:
        extra = k - q
        weeks = math.ceil(extra / w)
        last_week = extra - w * (weeks - 1)
        min_days = days_left + 7 * (weeks - 1) + last_week

    # ความเหนื่อยขั้นต่ำ: แคลที่เหลือต้องเผาอย่างน้อย rem
    min_strain = STRAIN_KCAL_W * rem / limits["max_kcal"]
    return min_days + min_strain


def strain_cost(session, prev_session, consec, limits):
    cost = STRAIN_KCAL_W * session["kcal"] / limits["max_kcal"]
    if prev_session is not None:
        cost += OVERLAP_W * len(session["muscle"] & prev_session["muscle"])
    cost += CONSEC_W * consec
    return cost


def a_star(target, sessions, limits, consec0=0, last0=-1, horizon=7):
    # ค้นหาแผน 1 สัปดาห์ (ไม่เกิน horizon วัน) ให้ได้แคลสะสมถึง target
    # consec0, last0 = สถานะที่ต่อมาจากสัปดาห์ก่อนหน้า
    max_session_kcal = max(s["kcal"] for s in sessions)
    start = (0, 0.0, consec0, 0, last0)
    h0 = heuristic(start, target, max_session_kcal, limits)
    queue = [(h0, h0, 0.0, start, [])]

    # visited: เก็บ "ชุดที่ดีที่สุด" (kcal, g) ของแต่ละสถานะหลัก (day, consec, week_count, last)
    # ถ้าสถานะใหม่เผาผลาญได้ไม่มากกว่า และต้นทุนไม่ต่ำกว่า สถานะที่มีอยู่ -> ตัดทิ้ง (dominated)
    visited = {}

    def push(g, n, route):
        key = (n[0], n[2], n[3], n[4])
        kcal = n[1]
        frontier = visited.setdefault(key, [])
        for k2, g2 in frontier:
            if k2 >= kcal and g2 <= g:
                return
        frontier[:] = [(k2, g2) for k2, g2 in frontier if not (kcal >= k2 and g <= g2)]
        frontier.append((kcal, g))
        f = g + heuristic(n, target, max_session_kcal, limits)
        heapq.heappush(queue, (f, f - g, g, n, route))

    visited[(0, consec0, 0, last0)] = [(0.0, 0.0)]

    while queue:
        f, h, g, n, route = heapq.heappop(queue)
        day, kcal, consec, week_count, last = n

        if kcal >= target:
            return route, g, kcal

        # ถ้าถูกสถานะอื่นที่ดีกว่าแทนที่ไปแล้ว ให้ข้าม
        if (kcal, g) not in visited[(day, consec, week_count, last)]:
            continue

        if day >= horizon:
            continue

        new_day = day + 1
        can_workout = consec < limits["max_consec"] and week_count < limits["max_days"]

        # ทางเลือกที่ 1: วันพัก (บังคับเมื่อออกครบโควตา, เลือกได้เมื่อเพิ่งออกกำลังกายมา)
        if not can_workout or consec > 0:
            new_week = 0 if new_day % 7 == 0 else week_count
            push(g + 1.0, (new_day, kcal, 0, new_week, -1), route + [-1])

        # ทางเลือกที่ 2: ออกกำลังกายด้วยชุดใดชุดหนึ่ง
        if can_workout:
            prev = sessions[last] if last >= 0 else None
            new_week = 0 if new_day % 7 == 0 else week_count + 1
            for i, sess in enumerate(sessions):
                nxt = (new_day, round(kcal + sess["kcal"], 1), consec + 1, new_week, i)
                push(g + 1.0 + strain_cost(sess, prev, consec, limits), nxt, route + [i])

    return None, 0, 0


def plan_weeks(target, sessions, limits):
    # A* วางแผนทีละสัปดาห์ แล้วต่อกัน
    # (ถ้าค้นหาทั้งแผนยาวทีเดียว จำนวนสถานะจะเพิ่มแบบทวีคูณและช้ามาก)
    best = max(s["kcal"] for s in sessions)
    weekly_cap = limits["max_days"] * best
    route, total_cost, done = [], 0.0, 0.0
    consec, last = 0, -1

    for _ in range(MAX_WEEKS):
        remaining = target - done
        # สัปดาห์สุดท้าย: ลองให้จบเลยถ้าทำได้ / สัปดาห์ทั่วไป: ตั้งเป้าไม่หนักเต็มเพดาน
        goals = []
        if remaining <= weekly_cap:
            goals.append(remaining)
        goals += [min(remaining, f * weekly_cap) for f in (WEEK_LOAD, 0.7, 0.5)]

        for goal in goals:
            week, g, kcal = a_star(goal, sessions, limits, consec, last)
            if week is not None:
                break
        else:
            return None, 0, 0

        done += kcal
        total_cost += g
        finished = done >= target
        if not finished:
            week = week + [-1] * (7 - len(week))     # เติมวันพักให้ครบสัปดาห์
        route += week

        # ส่งต่อสถานะไปสัปดาห์ถัดไป
        consec = 0
        for a in reversed(week):
            if a == -1:
                break
            consec += 1
        last = week[-1]

        if finished:
            return route, total_cost, done
    return None, 0, 0


# =====================================================================
#  6) รับข้อมูล + แสดงผล
# =====================================================================
def ask_number(text, lo, hi, cast=float):
    while True:
        try:
            v = cast(input(text))
            if lo <= v <= hi:
                return v
        except ValueError:
            pass
        print(f"  กรุณากรอกตัวเลขระหว่าง {lo} - {hi}")


def ask_diseases():
    names = list(DISEASE_MENU.keys())
    print("\nโรคประจำตัว/อาการ (เลือกได้หลายข้อ คั่นด้วยเครื่องหมายจุลภาค เช่น 1,4):")
    print("  0. ไม่มี")
    for i, d in enumerate(names, 1):
        print(f"  {i}. {d}")
    while True:
        raw = input("เลือก: ").replace(" ", "")
        if raw in ("", "0"):
            return []
        try:
            idx = {int(x) for x in raw.split(",")}
            if all(1 <= i <= len(names) for i in idx):
                return [names[i - 1] for i in sorted(idx)]
        except ValueError:
            pass
        print("  รูปแบบไม่ถูกต้อง")


def show_plan(route, total_cost, total_kcal, sessions, target):
    print("\n" + "=" * 60)
    print("แผนการออกกำลังกาย")
    print("=" * 60)
    done = 0.0
    for day, a in enumerate(route):
        if day % 7 == 0:
            print(f"\n--- สัปดาห์ที่ {day // 7 + 1} ---")
        if a == -1:
            print(f"วัน {day + 1:>3}: พัก")
            continue
        s = sessions[a]
        done += s["kcal"]
        print(f"วัน {day + 1:>3}: ออกกำลังกาย (~{s['kcal']:.1f} kcal | สะสม {done:.1f})")
        for e in s["exs"]:
            adj = "  [ลดเหลือ %d เซ็ต]" % e["sets"] if e["status"] == "modify" else ""
            print(f"          - {e['name']} ({e['muscle_text']}) "
                  f"{e['sets']} เซ็ต x {e['reps']}{adj}")
            print(f"            หมายเหตุจาก Excel: {e['note']}")

    workout_days = sum(1 for a in route if a != -1)
    print("\n" + "-" * 60)
    print(f"แคลที่ต้องการลด : {target:.0f} kcal")
    print(f"แคลที่ได้ตามแผน  : {total_kcal:.1f} kcal")
    print(f"ระยะเวลารวม      : {len(route)} วัน (ออกกำลังกาย {workout_days} วัน, พัก {len(route) - workout_days} วัน)")
    print(f"ต้นทุนรวม A* (g) : {total_cost:.2f}")


def main():
    print("=== ระบบแนะนำการออกกำลังกาย (A* Search) ===")
    age = ask_number("อายุ (ปี): ", 1, 120, int)
    weight = ask_number("น้ำหนัก (กก.): ", 10, 300)
    height = ask_number("ส่วนสูง (ซม.): ", 50, 250)
    diseases = ask_diseases()
    target = ask_number("แคลอรีที่ต้องการลด (kcal): ", 1, 100000)

    exercises = load_exercises()
    ages = [(e["age_lo"], e["age_hi"]) for e in exercises]
    if not any(lo <= age <= hi for lo, hi in ages):
        print(f"\nข้อมูลใน Excel รองรับอายุ {min(a[0] for a in ages)}-{max(a[1] for a in ages)} ปีเท่านั้น")
        return

    bmi = weight / (height / 100) ** 2
    print(f"\nอายุ {age} ปี | BMI = {bmi:.1f} | โรค: {', '.join(diseases) if diseases else 'ไม่มี'}")

    limits, notes = build_limits(age, bmi, diseases)
    for t in notes:
        print(" *", t)
    print(f" * ขีดจำกัด: ออกได้ไม่เกิน {limits['max_days']} วัน/สัปดาห์, "
          f"ติดกันไม่เกิน {limits['max_consec']} วัน, "
          f"ไม่เกิน {limits['max_kcal']:.0f} kcal/วัน")

    allowed, removed = filter_exercises(exercises, age, diseases)
    for name, why in removed:
        print(f" x ตัดท่า {name} ออก (ไม่เหมาะกับ: {', '.join(why)})")
    if not allowed:
        print("\nไม่พบท่าที่เหมาะสม กรุณาปรึกษาแพทย์หรือผู้เชี่ยวชาญก่อนออกกำลังกาย")
        return

    sessions = build_sessions(allowed, limits)

    # เช็กก่อนว่าเป้าหมายทำได้จริงภายใน MAX_WEEKS สัปดาห์หรือไม่
    best = max(x["kcal"] for x in sessions)
    weekly_cap = limits["max_days"] * best
    max_possible = (MAX_WEEKS - 1) * WEEK_LOAD * weekly_cap + weekly_cap   # สอดคล้องกับ plan_weeks
    if target > max_possible:
        print(f"\nเป้าหมายสูงเกินไปสำหรับการลดอย่างปลอดภัยภายใน {MAX_WEEKS} สัปดาห์ "
              f"(ทำได้สูงสุดประมาณ {max_possible:.0f} kcal) ลองลดเป้าหมายลง")
        return

    route, total_cost, total_kcal = plan_weeks(target, sessions, limits)

    if route:
        show_plan(route, total_cost, total_kcal, sessions, target)
        print("\nหมายเหตุ: เป็นคำแนะนำเบื้องต้น หากมีโรคประจำตัวควรปรึกษาแพทย์ก่อนเริ่มออกกำลังกาย")
    else:
        print(f"\nไม่พบแผนที่ปลอดภัยภายใน {MAX_WEEKS} สัปดาห์ ลองลดเป้าหมายแคลลง")


if __name__ == "__main__":
    main()
