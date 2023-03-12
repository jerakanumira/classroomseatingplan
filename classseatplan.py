import tkinter as tk
import openpyxl

class ClassroomSeatingPlanApp:
    def __init__(self, master):
        self.master = master
        master.title("Classroom Seating Plan")

        # add title label
        title_label = tk.Label(master, text="Classroom Seating Plan", font=("Arial", 20))
        title_label.grid(row=0, column=0, columnspan=2, pady=10)

        # add row and column labels
        row_label = tk.Label(master, text="Number of rows:")
        row_label.grid(row=1, column=0, padx=5, pady=5)
        self.row_entry = tk.Entry(master)
        self.row_entry.grid(row=1, column=1, padx=5, pady=5)
        col_label = tk.Label(master, text="Number of columns:")
        col_label.grid(row=2, column=0, padx=5, pady=5)
        self.col_entry = tk.Entry(master)
        self.col_entry.grid(row=2, column=1, padx=5, pady=5)

        # add submit button
        submit_button = tk.Button(master, text="Generate Seating Plan", command=self.generate_seating_plan)
        submit_button.grid(row=3, column=0, columnspan=2, pady=10)

    def generate_seating_plan(self):
        # get classroom size from entry widgets
        num_rows = int(self.row_entry.get())
        num_cols = int(self.col_entry.get())

        # load student data from Excel document
        workbook = openpyxl.load_workbook("filepath")
        worksheet = workbook.active
        data = []
        for row in worksheet.iter_rows(min_row=2, values_only=True):
            data.append(row[:3])
        print(f"Loaded {len(data)} students from Excel")

        # sort students by NAPLAN score
        data = sorted(data, key=lambda x: x[2] or 0)
        print("Students sorted by NAPLAN score")

        # create seating plan
        plan = [[None for j in range(num_cols)] for i in range(num_rows)]
        prev_gender = None
        prev_student = None  # define prev_student to None
        for i in range(num_rows):
            for j in range(num_cols):
                # find next available student with matching gender
                prev_gender = prev_student[1] if prev_student is not None else None
                found_student = False
                for s in data:
                    if s[1] == "M" and (j == 0 or prev_gender != "F"):
                        continue
                    if s[1] == "F" and (j == 0 or prev_gender != "M"):
                        continue
                    if s not in [data[p] for row in plan for p in row if p is not None]:
                        plan[i][j] = data.index(s)
                        prev_student = s  # update prev_student
                        found_student = True
                        break
                # handle case where no student was found
                if not found_student:
                    prev_student = None
                    break
            # handle case where no student was found
            if not prev_student:
                break

        print("Seating plan created")
        print("data:", data)
        print("plan:", plan)

        # create new window for seating plan display
        plan_window = tk.Toplevel(self.master)

        # create labels for seating plan
        labels = []
        for i in range(num_rows):
            for j in range(num_cols):
                name = ""
                if plan[i][j] is not None:
                    if plan[i][j] < len(data) and data[plan[i][j]]:
                        name = data[plan[i][j]][0]
                label = tk.Label(plan_window, text=name, width=20)
                label.grid(row=i, column=j, padx=5, pady=5)
                labels.append(label)

        # apply padding to labels
        for i, label in enumerate(labels):
            if (i % 2 == 1) and ((i // num_cols) % 2 == 1):
                label.grid(row=i // num_cols, column=i % num_cols, padx=40, pady=5)
            else:
                label.grid(row=i // num_cols, column=i % num_cols, padx=5, pady=5)

        print("Seating plan displayed")

# create and run the app
root = tk.Tk()
app = ClassroomSeatingPlanApp(root)
root.mainloop()
