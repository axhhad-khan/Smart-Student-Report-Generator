import pandas as pd 
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.filedialog import asksaveasfilename, askopenfilename

# A list to store user data
users = []

# Function to validate int input
def validate_input(value, input_type, min_val=None, max_val=None):
    if value.strip() == "":
        return False
    if input_type == int:
        if value.isdigit() and (min_val is None or min_val <= int(value) <= max_val):
            return int(value)
    if input_type == float:
        if value.replace('.', '', 1).isdigit():
            val = float(value)
            if min_val is None or min_val <= val <= max_val:
                return val
    return False

# Function to validate str input
def validate_name(name):
    return name.isalpha()

# Function to check if a user already exists
def user_exists(name, age):
    for user in users:
        if user["name"].lower() == name.lower() and user["age"] == age:
            return True
    return False

# Function to add a new user
def add_user():
    name = name_entry.get().strip()
    if not validate_name(name):
        messagebox.showerror("Error", "Name must contain only alphabetic characters.")
        return
    age = validate_input(age_entry.get(), int, 15, 25)
    marks1 = validate_input(marks1_entry.get(), float, 0, 100)
    marks2 = validate_input(marks2_entry.get(), float, 0, 100)
    marks3 = validate_input(marks3_entry.get(), float, 0, 100)

    if not name:
        messagebox.showerror("Error", "Name cannot be empty.")
        return
    if not age:
        messagebox.showerror("Error", "Enter a valid age (15-25).")
        return
    if not marks1 or not marks2 or not marks3:
        messagebox.showerror("Error", "Marks must be between 0 and 100.")
        return

    if user_exists(name, age):
        messagebox.showerror("Error", "User already exists.")
        return

    total = marks1 + marks2 + marks3
    percentage = (total / 300) * 100
    discount = "Yes" if age <= 18 or percentage >= 90 else "No"

    user = {
        "name": name,
        "age": age,
        "subject1": marks1,
        "subject2": marks2,
        "subject3": marks3,
        "total_marks": total,
        "percentage": percentage,
        "discount": discount
    }
    users.append(user)
    messagebox.showinfo("Success", "User added successfully!")
    clear_entries()


# Function to clear input fields
def clear_entries():
    name_entry.delete(0, tk.END)
    age_entry.delete(0, tk.END)
    marks1_entry.delete(0, tk.END)
    marks2_entry.delete(0, tk.END)
    marks3_entry.delete(0, tk.END)

# Function to search for a user
def search_user():
    search_name = search_entry.get().strip()
    if not search_name:
        messagebox.showerror("Error", "Enter a name to search.")
        return

    for user in users:
        if user["name"].lower() == search_name.lower():
            result = (
                f"Name: {user['name']}\n"
                f"Age: {user['age']}\n"
                f"Subject 1: {user['subject1']}\n"
                f"Subject 2: {user['subject2']}\n"
                f"Subject 3: {user['subject3']}\n"
                f"Total Marks: {user['total_marks']}\n"
                f"Percentage: {user['percentage']:.2f}\n"
                f"Discount: {user['discount']}"
            )
            messagebox.showinfo("User Found", result)
            return

    messagebox.showerror("Error", "User not found.")

# Function to sort users by a given key
def sort_users(key):
    if not users:
        messagebox.showerror("Error", "No users to sort.")
        return

    sorted_users = sorted(users, key=lambda x: x[key], reverse=(key in ["total_marks", "percentage"]))
    display_text = "\n".join([f"{user['name']} - {key.capitalize()}: {user[key]}" for user in sorted_users])
    messagebox.showinfo("Sorted Users", display_text)

# Function to save data to Excel
def save_to_excel():
    if not users:
        messagebox.showerror("Error", "No users to save.")
        return

    file_path = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        data = pd.DataFrame(users)
        data.to_excel(file_path, index=False)
        messagebox.showinfo("Success", "Data saved successfully!")

# Function to load data from Excel
def load_from_excel():
    file_path = askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        data = pd.read_excel(file_path)
        for _, row in data.iterrows():
            name = row["name"]
            age = row["age"]
            if not user_exists(name, age):
                user = {
                    "name": name,
                    "age": age,
                    "subject1": row["subject1"],
                    "subject2": row["subject2"],
                    "subject3": row["subject3"],
                    "total_marks": row["total_marks"],
                    "percentage": row["percentage"],
                    "discount": row["discount"]
                }
                users.append(user)
        messagebox.showinfo("Success", "Data loaded successfully!")

# Main window
app = tk.Tk()
app.title("Smart Student Report Generator")
app.geometry("550x700")
app.resizable(False, False)

# Header
header_label = tk.Label(app, text="Smart Student Report Generator", font=("Arial", 20, "bold"), pady=10)
header_label.pack()

# Input section
input_frame = tk.LabelFrame(app, text="Input Details", padx=10, pady=10, font=("Arial", 12))
input_frame.pack(fill=tk.X, padx=20, pady=10)

ttk.Label(input_frame, text="Name:").grid(row=0, column=0, sticky=tk.W, pady=5)
name_entry = ttk.Entry(input_frame, width=30)
name_entry.grid(row=0, column=1, pady=5)

ttk.Label(input_frame, text="Age:").grid(row=1, column=0, sticky=tk.W, pady=5)
age_entry = ttk.Entry(input_frame, width=30)
age_entry.grid(row=1, column=1, pady=5)

ttk.Label(input_frame, text="Subject 1(out of 100):").grid(row=2, column=0, sticky=tk.W, pady=5)
marks1_entry = ttk.Entry(input_frame, width=30)
marks1_entry.grid(row=2, column=1, pady=5)

ttk.Label(input_frame, text="Subject 2(out of 100):").grid(row=3, column=0, sticky=tk.W, pady=5)
marks2_entry = ttk.Entry(input_frame, width=30)
marks2_entry.grid(row=3, column=1, pady=5)

ttk.Label(input_frame, text="Subject 3(out of 100):").grid(row=4, column=0, sticky=tk.W, pady=5)
marks3_entry = ttk.Entry(input_frame, width=30)
marks3_entry.grid(row=4, column=1, pady=5)

add_button = ttk.Button(input_frame, text="Add User", command=add_user)
add_button.grid(row=5, column=0, columnspan=2, pady=10)

# Sorting section
sort_frame = tk.LabelFrame(app, text="Sorting Options", padx=10, pady=10, font=("Arial", 12))
sort_frame.pack(fill=tk.X, padx=20, pady=10)

sort_age_button = ttk.Button(sort_frame, text="Sort by Age", command=lambda: sort_users("age"))
sort_age_button.grid(row=0, column=0, padx=10, pady=5)

sort_total_marks_button = ttk.Button(sort_frame, text="Sort by Total Marks", command=lambda: sort_users("total_marks"))
sort_total_marks_button.grid(row=0, column=1, padx=10, pady=5)

sort_percentage_button = ttk.Button(sort_frame, text="Sort by Percentage", command=lambda: sort_users("percentage"))
sort_percentage_button.grid(row=0, column=2, padx=10, pady=5)


# Search section
search_frame = tk.LabelFrame(app, text="Search User", padx=10, pady=10, font=("Arial", 12))
search_frame.pack(fill=tk.X, padx=20, pady=10)

ttk.Label(search_frame, text="Enter Name:").grid(row=0, column=0, padx=10, pady=5)
search_entry = ttk.Entry(search_frame, width=30)
search_entry.grid(row=0, column=1, padx=10, pady=5)

search_button = ttk.Button(search_frame, text="Search", command=search_user)
search_button.grid(row=0, column=2, padx=10, pady=5)

sort_frame.pack(fill=tk.X, padx=20, pady=10)


# File actions
file_frame = tk.LabelFrame(app, text="File Actions", padx=10, pady=10, font=("Arial", 12))
file_frame.pack(fill=tk.X, padx=20, pady=10)

save_button = ttk.Button(file_frame, text="Save to Excel", command=save_to_excel)
save_button.grid(row=0, column=0, padx=10, pady=5)

load_button = ttk.Button(file_frame, text="Load from Excel", command=load_from_excel)
load_button.grid(row=0, column=1, padx=10, pady=5)

exit_button = ttk.Button(file_frame, text="Exit", command=app.quit)
exit_button.grid(row=0, column=2, padx=10, pady=5)


# Run the application
app.mainloop()