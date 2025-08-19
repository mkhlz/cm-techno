import pandas as pd
import random
from faker import Faker

fake = Faker()

# --- Load your base classes list ---
classes_df = pd.read_excel("Classes_Data.xlsx")

# Add realistic teacher names
classes_df["Teacher Name"] = [fake.name() for _ in range(len(classes_df))]

# Random schedules (weekday and time)
weekdays = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
times = ["9:00 AM", "11:00 AM", "2:00 PM", "4:00 PM"]
classes_df["Schedule"] = [f"{random.choice(weekdays)} {random.choice(times)}" for _ in range(len(classes_df))]

# Assign rooms
classes_df["Room"] = [f"R{random.randint(101, 120)}" for _ in range(len(classes_df))]

# Add capacity and enrollment
classes_df["Capacity"] = [random.randint(15, 40) for _ in range(len(classes_df))]
classes_df["Enrollment"] = [random.randint(10, cap) for cap in classes_df["Capacity"]]

# Term (academic year)
classes_df["Term"] = ["2025"] * len(classes_df)

# --- Generate Fees data ---
num_students = 5000
fees_data = {
    "Student ID": [f"S{str(i).zfill(4)}" for i in range(1, num_students + 1)],
    "Name": [fake.first_name() for _ in range(num_students)],
    "Admission Date": [fake.date_between(start_date="-2y", end_date="today") for _ in range(num_students)],
    "Fee Amount": [random.choice([15000, 20000, 25000]) for _ in range(num_students)],
    "Payment Status": [random.choice(["Paid", "Partial", "Pending"]) for _ in range(num_students)],
    "Payment Date": [None] * num_students,
    "Payment Mode": [None] * num_students,
    "Transaction ID": [None] * num_students
}

# Fill payment details where applicable
for i in range(num_students):
    if fees_data["Payment Status"][i] != "Pending":
        fees_data["Payment Date"][i] = fake.date_between(
            start_date=fees_data["Admission Date"][i], end_date="today"
        )
        fees_data["Payment Mode"][i] = random.choice(["UPI", "Cash", "Card", "Bank Transfer"])
        fees_data["Transaction ID"][i] = "TXN" + str(random.randint(1000, 9999))

fees_df = pd.DataFrame(fees_data)

# --- Generate Attendance data ---
num_attendance_records = 10000
attendance_statuses = ["Present", "Absent", "Late"]
absence_reasons = ["Sick", "Family Emergency", "Travel", ""]

attendance_data = []
for _ in range(num_attendance_records):
    student = fees_df.sample(1).iloc[0]
    course = classes_df.sample(1).iloc[0]
    status = random.choice(attendance_statuses)
    reason = random.choice(absence_reasons) if status == "Absent" else ""
    
    attendance_data.append({
        "Date": fake.date_between(start_date="-90d", end_date="today"),
        "Student ID": student["Student ID"],
        "Name": student["Name"],
        "Class ID": course["Class ID"],
        "Attendance Status": status,
        "Reason": reason
    })

attendance_df = pd.DataFrame(attendance_data)

# --- Save all sheets into one workbook ---
output_path = "School_Management_Full.xlsx"
with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
    classes_df.to_excel(writer, sheet_name="Classes", index=False)
    fees_df.to_excel(writer, sheet_name="Fees", index=False)
    attendance_df.to_excel(writer, sheet_name="Attendance", index=False)

print(f"✅ Realistic school management workbook created: {output_path}")
