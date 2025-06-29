import tkinter as tk
from tkinter import messagebox, ttk, simpledialog
import os
import csv
from datetime import datetime, date, timedelta

class VehicleManager:
    """
    Manages vehicle check-out, check-in, and history records using a Tkinter GUI.
    Stores vehicle information in 'vehicle_info.csv' and historical data in 'vehicle_history.csv'.
    """
    def __init__(self, root):
        """
        Initializes the VehicleManager application.

        Args:
            root: The root Tkinter window.
        """
        self.root = root
        self.root.title("Vehicle Manager")
        self.root.geometry("1100x600")

        self.init_ui()
        self.update_table()

        # Initialize service file if it doesn't exist
        self.init_service_file()

    def init_service_file(self):
        """
        Ensures the vehicle_services.csv file exists with the correct header.
        This file stores the configuration for service items for each vehicle.
        Header: Vehicle,Service Item,Mileage Interval (miles),Time Interval (days),Last Service Date (YYYY-MM-DD),Last Service Mileage
        """
        file_name = "vehicle_services.csv"
        header = ["Vehicle", "Service Item", "Mileage Interval (miles)", "Time Interval (days)", "Last Service Date (YYYY-MM-DD)", "Last Service Mileage"]
        if not os.path.exists(file_name) or os.path.getsize(file_name) == 0:
            with open(file_name, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(header)

    def init_ui(self):
        """
        Initializes the user interface elements and their layout.
        Divides the main window into left (input/controls) and right (table display) frames.
        """
        # --- Frame Setup ---
        self.left_frame = tk.Frame(self.root)
        self.left_frame.grid(row=0, column=0, sticky='nsew', padx=10, pady=10)
        self.right_frame = tk.Frame(self.root)
        self.right_frame.grid(row=0, column=1, sticky='nsew', padx=10, pady=10)

        # Configure grid weights to make frames resize with window
        self.root.grid_columnconfigure(0, weight=1, uniform="group1")
        self.root.grid_columnconfigure(1, weight=2, uniform="group1")
        self.root.grid_rowconfigure(0, weight=1)

        # --- Left Frame: Input Fields ---
        labels = ["Vehicle:", "Check out Purpose:", "User:", "Check Out (YYYY-MM-DD):", "Estimated Check In (YYYY-MM-DD):"]
        self.entries = {}

        for i, text in enumerate(labels):
            tk.Label(self.left_frame, text=text).grid(row=i, column=0, sticky='w', pady=5)
            entry = tk.Entry(self.left_frame)
            entry.grid(row=i, column=1, sticky='ew', pady=5, padx=(5,0))
            self.entries[text] = entry
        self.left_frame.grid_columnconfigure(1, weight=1)

        self.save_button = tk.Button(self.left_frame, text="Save", command=self.save_info)
        self.save_button.grid(row=5, column=0, columnspan=2, pady=10)

        self.info_display = tk.Text(self.left_frame, height=10, width=50)
        self.info_display.grid(row=6, column=0, columnspan=2, pady=10, sticky='nsew')
        self.left_frame.grid_rowconfigure(6, weight=1)

        # Buttons at the bottom left
        btn_bottom_left_frame = tk.Frame(self.left_frame)
        btn_bottom_left_frame.grid(row=7, column=0, columnspan=2, pady=10, sticky='w')
        tk.Button(btn_bottom_left_frame, text="Check In Vehicle", command=self.open_checkin_window).grid(row=0, column=0, padx=5)
        tk.Button(btn_bottom_left_frame, text="View Vehicle History", command=self.view_history).grid(row=0, column=1, padx=5)
        
        # New button for Vehicle Service Tracker
        tk.Button(btn_bottom_left_frame, text="Track Vehicle Services", command=self.open_service_tracker_window).grid(row=0, column=2, padx=5)


        # --- Right Frame: Table and Delete Buttons ---
        btn_delete_frame = tk.Frame(self.right_frame)
        btn_delete_frame.grid(row=0, column=0, sticky='w', pady=(0,10))
        tk.Button(btn_delete_frame, text="Delete Selected", command=self.delete_selected).grid(row=0, column=0, padx=5)
        tk.Button(btn_delete_frame, text="Delete All", command=self.delete_all).grid(row=0, column=1, padx=5)

        self.table_frame = tk.Frame(self.right_frame)
        self.table_frame.grid(row=1, column=0, sticky='nsew')
        self.right_frame.grid_rowconfigure(1, weight=1)
        self.right_frame.grid_columnconfigure(0, weight=1)

        # Treeview (table) setup
        self.table = ttk.Treeview(self.table_frame, columns=[
            "Vehicle", "Purpose", "User", "Checked Out", "Estimated Check In",
            "Status", "Actual Check In", "Fuel (%)", "Comments", "Mileage at Check In"], show='headings', selectmode='browse')
        
        # Configure column headings and widths
        for col in self.table["columns"]:
            self.table.heading(col, text=col)
            self.table.column(col, width=120, anchor='center')
        
        # Define tags for row background colors based on status
        self.table.tag_configure('active', background='')
        self.table.tag_configure('inactive', background='pale green')
        self.table.tag_configure('overdue', background='yellow')

        # Add scrollbar to the table
        vsb = ttk.Scrollbar(self.table_frame, orient="vertical", command=self.table.yview)
        self.table.configure(yscrollcommand=vsb.set)
        
        self.table.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        self.table_frame.grid_rowconfigure(0, weight=1)
        self.table_frame.grid_columnconfigure(0, weight=1)
        self.table.bind("<<TreeviewSelect>>", self.on_table_select)

    def try_parse_date(self, date_str):
        """
        Attempts to parse a date string into a datetime.date object using multiple formats.

        Args:
            date_str (str): The date string to parse.

        Returns:
            datetime.date or None: The parsed date object if successful, otherwise None.
        """
        formats = ("%Y-%m-%d", "%m/%d/%Y", "%d-%m-%Y", "%d/%m/%Y")
        for fmt in formats:
            try:
                return datetime.strptime(date_str, fmt).date()
            except ValueError:
                continue
        return None

    def save_info(self):
        """
        Saves new vehicle check-out information or updates existing vehicle status.
        Includes validation for input fields and date logic.
        """
        # Get data from input fields
        vehicle = self.entries["Vehicle:"].get().strip()
        purpose = self.entries["Check out Purpose:"].get().strip()
        user = self.entries["User:"].get().strip()
        checked_out_str = self.entries["Check Out (YYYY-MM-DD):"].get().strip()
        checked_in_str = self.entries["Estimated Check In (YYYY-MM-DD):"].get().strip()

        # Input validation: Check if all fields are filled
        if not all([vehicle, purpose, user, checked_out_str, checked_in_str]):
            messagebox.showwarning("Input Error", "Please fill in all fields")
            return

        # Date parsing and validation
        checked_out = self.try_parse_date(checked_out_str)
        estimated_check_in = self.try_parse_date(checked_in_str)

        if not checked_out or not estimated_check_in:
            messagebox.showwarning("Date Error", "Invalid date format for Check Out or Estimated Check In. Please use APAC-MM-DD.")
            return

        # Rule 1: Checkout date cannot be after Estimated Check In
        if checked_out > estimated_check_in:
            messagebox.showwarning("Date Rule Violation", "Check Out date cannot be after Estimated Check In date.")
            return

        # Save or update the vehicle info in the CSV.
        # For a new checkout, actual_checkin, fuel, comments, and mileage at check-in are None.
        self.save_or_update_csv(vehicle, purpose, user, checked_out.strftime("%Y-%m-%d"), estimated_check_in.strftime("%Y-%m-%d"), None, None, None, None)
        
        # Display saved information in the info_display text box
        self.info_display.insert(tk.END, f"Vehicle: {vehicle}\nPurpose: {purpose}\nUser: {user}\nChecked Out: {checked_out_str}\nEstimated Check In: {checked_in_str}\n\n")
        self.info_display.see(tk.END)
        self.clear_entries()
        self.update_table()

    def save_or_update_csv(self, vehicle, purpose, user, checked_out, estimated_check_in, actual_checkin, fuel, comments, mileage_at_checkin):
        """
        Saves or updates a vehicle's record in 'vehicle_info.csv'.
        Determines the status (ACTIVE, INACTIVE, OVERDUE) based on dates.

        Args:
            vehicle (str): Vehicle identifier.
            purpose (str): Purpose of check-out.
            user (str): User who checked out the vehicle.
            checked_out (str): Check-out date (YYYY-MM-DD).
            estimated_check_in (str): Estimated check-in date (YYYY-MM-DD).
            actual_checkin (str or None): Actual check-in date (YYYY-MM-DD) if checked in, else None.
            fuel (str or None): Fuel level at check-in, else None.
            comments (str or None): Check-in comments, else None.
            mileage_at_checkin (str or None): Mileage at check-in, else None.
        """
        file_name = "vehicle_info.csv"
        header = ["Vehicle", "Purpose", "User", "Checked Out", "Estimated Check In", "Status", "Actual Check In", "Fuel (%)", "Comments", "Mileage at Check In"]
        today = date.today()

        # Determine the status of the vehicle
        status = "INACTIVE"
        if actual_checkin is None:
            estimated_check_in_date_obj = self.try_parse_date(estimated_check_in)
            if estimated_check_in_date_obj:
                if today > estimated_check_in_date_obj:
                    status = "OVERDUE"
                else:
                    status = "ACTIVE"

        data = []
        if os.path.exists(file_name):
            with open(file_name, 'r', newline='') as f:
                reader = csv.reader(f)
                data = list(reader)

            if not data or data[0] != header:
                data = [header]
        else:
            data = [header]

        found = False
        new_row = [vehicle, purpose, user, checked_out, estimated_check_in, status, actual_checkin or "", fuel or "", comments or "", mileage_at_checkin or ""]

        for i in range(1, len(data)):
            if data[i][0] == vehicle:
                data[i] = new_row
                found = True
                break
        
        if not found:
            data.append(new_row)

        with open(file_name, 'w', newline='') as f:
            csv.writer(f).writerows(data)

    def clear_entries(self):
        """Clears the text in all input Entry widgets."""
        for entry in self.entries.values():
            entry.delete(0, tk.END)

    def update_table(self):
        """
        Refreshes the data displayed in the main Treeview table.
        Loads data from 'vehicle_info.csv' and applies appropriate styling (colors).
        """
        file_name = "vehicle_info.csv"
        for row in self.table.get_children():
            self.table.delete(row)

        if os.path.exists(file_name):
            with open(file_name, 'r', newline='') as f:
                reader = csv.reader(f)
                data = list(reader)
                if not data or len(data) < 2:
                    return

                for row_data in data[1:]:
                    status_col_index = 5
                    tag = 'inactive'
                    
                    if row_data[status_col_index].strip().upper() == "ACTIVE":
                        tag = 'active'
                    elif row_data[status_col_index].strip().upper() == "OVERDUE":
                        tag = 'overdue'

                    while len(row_data) < len(self.table["columns"]):
                        row_data.append("")

                    self.table.insert('', tk.END, values=row_data, tags=(tag,))

    def on_table_select(self, event):
        """
        Event handler for when a row in the main Treeview table is selected.
        Populates the input fields with the selected vehicle's data.
        """
        selected = self.table.selection()
        if selected:
            values = self.table.item(selected[0], 'values')
            if values:
                self.clear_entries()
                self.entries["Vehicle:"].insert(0, values[0])
                if values[1]: self.entries["Check out Purpose:"].insert(0, values[1])
                if values[2]: self.entries["User:"].insert(0, values[2])
                if values[3]: self.entries["Check Out (YYYY-MM-DD):"].insert(0, values[3])
                if values[4]: self.entries["Estimated Check In (YYYY-MM-DD):"].insert(0, values[4])

    def delete_selected(self):
        """
        Deletes the selected vehicle record from 'vehicle_info.csv' after confirmation.
        Also removes associated service records from 'vehicle_services.csv'.
        """
        selected = self.table.selection()
        if not selected:
            messagebox.showwarning("Selection Error", "No vehicle selected")
            return
        
        confirm = messagebox.askyesno("Confirm Deletion", "Are you sure you want to delete the selected vehicle from the main list? This will also delete its service records.")
        if not confirm:
            return

        vehicle_to_delete = self.table.item(selected[0], 'values')[0]
        
        file_name = "vehicle_info.csv"
        if os.path.exists(file_name):
            with open(file_name, 'r', newline='') as f:
                data = list(csv.reader(f))
            
            new_data = [data[0]] + [row for row in data[1:] if row[0] != vehicle_to_delete]
            
            with open(file_name, 'w', newline='') as f:
                csv.writer(f).writerows(new_data)
        
        service_file_name = "vehicle_services.csv"
        if os.path.exists(service_file_name):
            with open(service_file_name, 'r', newline='') as f:
                data = list(csv.reader(f))
            
            if data and len(data) > 0:
                new_service_data = [data[0]] + [row for row in data[1:] if len(row) > 0 and row[0] != vehicle_to_delete]
            else:
                new_service_data = [["Vehicle", "Service Item", "Mileage Interval (miles)", "Time Interval (days)", "Last Service Date (YYYY-MM-DD)", "Last Service Mileage"]]
            
            with open(service_file_name, 'w', newline='') as f:
                csv.writer(f).writerows(new_service_data)

        self.clear_entries()
        self.update_table()
        messagebox.showinfo("Success", f"Vehicle '{vehicle_to_delete}' and its service records deleted.")

    def delete_all(self):
        """
        Deletes all records from 'vehicle_info.csv', 'vehicle_history.csv', and 'vehicle_services.csv' after double confirmation.
        """
        file_name = "vehicle_info.csv"
        history_file_name = "vehicle_history.csv"
        service_file_name = "vehicle_services.csv"
        
        confirm = messagebox.askyesno("Confirm Delete All", "Are you sure you want to delete ALL vehicle records, history, and service data? This cannot be undone.")
        if confirm:
            confirm_again = messagebox.askyesno("REALLY Delete All?", "This will PERMANENTLY delete ALL data. Are you absolutely sure?")
            if confirm_again:
                if os.path.exists(file_name):
                    os.remove(file_name)
                if os.path.exists(history_file_name):
                    os.remove(history_file_name)
                if os.path.exists(service_file_name):
                    os.remove(service_file_name)
                
                self.init_service_file()

                self.clear_entries()
                self.update_table()
                messagebox.showinfo("Success", "All vehicle records, history, and service data deleted.")

    def open_checkin_window(self):
        """
        Opens a new Toplevel window for vehicle check-in.
        Populates a dropdown with currently active (checked-out) vehicles.
        """
        self.checkin_win = tk.Toplevel(self.root)
        self.checkin_win.title("Vehicle Check In")

        tk.Label(self.checkin_win, text="Select Vehicle to Check In:").grid(row=0, column=0, padx=10, pady=5, sticky='w')
        self.active_vehicles = self.get_active_vehicles()
        
        if not self.active_vehicles:
            messagebox.showinfo("No Active Vehicles", "There are no vehicles currently checked out.")
            self.checkin_win.destroy()
            return
        
        self.vehicle_var = tk.StringVar()
        self.vehicle_dropdown = ttk.Combobox(self.checkin_win, values=self.active_vehicles, textvariable=self.vehicle_var, state='readonly')
        self.vehicle_dropdown.grid(row=0, column=1, padx=10, pady=5, sticky='ew')
        
        if self.active_vehicles:
            self.vehicle_dropdown.set(self.active_vehicles[0])

        tk.Label(self.checkin_win, text="Actual Check In (YYYY-MM-DD):").grid(row=1, column=0, padx=10, pady=5, sticky='w')
        self.actual_checkin_entry = tk.Entry(self.checkin_win)
        self.actual_checkin_entry.grid(row=1, column=1, padx=10, pady=5, sticky='ew')
        self.actual_checkin_entry.insert(0, date.today().strftime("%Y-%m-%d"))

        tk.Label(self.checkin_win, text="Fuel Percentage (%):").grid(row=2, column=0, padx=10, pady=5, sticky='w')
        self.fuel_entry = tk.Entry(self.checkin_win)
        self.fuel_entry.grid(row=2, column=1, padx=10, pady=5, sticky='ew')

        tk.Label(self.checkin_win, text="Comments:").grid(row=3, column=0, padx=10, pady=5, sticky='w')
        self.comments_entry = tk.Entry(self.checkin_win)
        self.comments_entry.grid(row=3, column=1, padx=10, pady=5, sticky='ew')

        tk.Label(self.checkin_win, text="Mileage at Check In:").grid(row=4, column=0, padx=10, pady=5, sticky='w')
        self.mileage_checkin_entry = tk.Entry(self.checkin_win)
        self.mileage_checkin_entry.grid(row=4, column=1, padx=10, pady=5, sticky='ew')


        tk.Button(self.checkin_win, text="Check In", command=self.check_in_vehicle).grid(row=5, column=0, columnspan=2, pady=10)
        self.checkin_win.grid_columnconfigure(1, weight=1)


    def get_active_vehicles(self):
        """
        Retrieves a list of vehicles that are currently "ACTIVE" (checked out)
        and have no actual check-in date recorded in 'vehicle_info.csv'.

        Returns:
            list: A list of vehicle IDs.
        """
        vehicles = []
        file_name = "vehicle_info.csv"
        if os.path.exists(file_name):
            with open(file_name, 'r', newline='') as file:
                reader = csv.DictReader(file)
                
                expected_header = ["Vehicle", "Purpose", "User", "Checked Out", "Estimated Check In", "Status", "Actual Check In", "Fuel (%)", "Comments", "Mileage at Check In"]
                try:
                    first_row = next(reader)
                    from collections import deque
                    data_rows = deque([first_row])
                    data_rows.extend(reader)
                    
                    if reader.fieldnames != expected_header and len(reader.fieldnames) == len(expected_header):
                        pass
                except StopIteration:
                    return []
                
                file.seek(0) 
                reader = csv.DictReader(file)
                next(reader, None)

                for row in reader:
                    if "Actual Check In" in row and not row["Actual Check In"].strip(): 
                        vehicles.append(row["Vehicle"])
        return vehicles

    def check_in_vehicle(self):
        """
        Handles the vehicle check-in process.
        Updates 'vehicle_info.csv' and archives the complete record in 'vehicle_history.csv'.
        Includes validation for actual check-in date against check-out date.
        """
        vehicle = self.vehicle_var.get()
        actual_checkin_str = self.actual_checkin_entry.get().strip()
        fuel_str = self.fuel_entry.get().strip()
        comments = self.comments_entry.get().strip()
        mileage_at_checkin_str = self.mileage_checkin_entry.get().strip()

        if not vehicle:
            messagebox.showwarning("Selection Error", "Please select a vehicle to check in.")
            return

        if not mileage_at_checkin_str:
            messagebox.showwarning("Input Error", "Please enter mileage at check in.")
            return

        try:
            mileage_at_checkin = int(mileage_at_checkin_str)
            if mileage_at_checkin < 0:
                raise ValueError("Mileage cannot be negative.")
        except ValueError:
            messagebox.showwarning("Input Error", "Mileage at Check In must be a valid positive number.")
            return

        actual_checkin_date = self.try_parse_date(actual_checkin_str)
        if not actual_checkin_date:
            messagebox.showwarning("Date Format Error", "Invalid actual check-in date format. Please use YYYY-MM-DD.")
            return

        fuel = ""
        if fuel_str:
            try:
                fuel = float(fuel_str)
                if not (0 <= fuel <= 100):
                    messagebox.showwarning("Fuel Error", "Fuel percentage must be a number between 0 and 100.")
                    return
            except ValueError:
                messagebox.showwarning("Fuel Error", "Fuel percentage must be a valid number.")
                return

        file_name = "vehicle_info.csv"
        history_file = "vehicle_history.csv"
        
        updated_main_data = []
        archived_row_details = None
        checkout_date_for_validation = None

        if os.path.exists(file_name):
            with open(file_name, 'r', newline='') as f:
                reader = csv.reader(f)
                header = next(reader)
                updated_main_data.append(header)

                for row in reader:
                    if row[0] == vehicle:
                        checkout_date_for_validation = self.try_parse_date(row[3]) 
                        
                        if checkout_date_for_validation and actual_checkin_date < checkout_date_for_validation:
                            messagebox.showwarning("Date Rule Violation", "Actual Check In date cannot be before the Check Out date.")
                            self.checkin_win.destroy()
                            return

                        original_checkout_mileage_str = row[9] if len(row) > 9 else ""
                        if original_checkout_mileage_str:
                            try:
                                original_checkout_mileage = int(original_checkout_mileage_str)
                                if mileage_at_checkin < original_checkout_mileage:
                                    messagebox.showwarning("Mileage Error", "Mileage at Check In cannot be less than Mileage at Check Out.")
                                    self.checkin_win.destroy()
                                    return
                            except ValueError:
                                pass


                        archived_row_details = list(row) 
                        
                        while len(archived_row_details) < 10:
                            archived_row_details.append("")

                        archived_row_details[5] = "INACTIVE"
                        archived_row_details[6] = actual_checkin_date.strftime("%Y-%m-%d")
                        archived_row_details[7] = str(fuel)
                        archived_row_details[8] = comments
                        archived_row_details[9] = str(mileage_at_checkin)


                        row[1] = ""
                        row[2] = ""
                        row[3] = ""
                        row[4] = ""
                        row[5] = "INACTIVE"
                        row[6] = ""
                        row[7] = ""
                        row[8] = ""
                        row[9] = str(mileage_at_checkin)
                        
                        updated_main_data.append(row)
                    else:
                        while len(row) < 10:
                            row.append("")
                        updated_main_data.append(row)

            if archived_row_details:
                with open(history_file, 'a', newline='') as hist_f:
                    writer = csv.writer(hist_f)
                    if not os.path.exists(history_file) or os.path.getsize(history_file) == 0:
                        hist_writer_header = ["Vehicle", "Purpose", "User", "Checked Out", "Estimated Check In", "Status", "Actual Check In", "Fuel (%)", "Comments", "Mileage at Check In"]
                        writer.writerow(hist_writer_header) 
                    writer.writerow(archived_row_details)

                with open(file_name, 'w', newline='') as f:
                    csv.writer(f).writerows(updated_main_data)

                messagebox.showinfo("Check In Success", f"Vehicle '{vehicle}' checked in successfully!")
            else:
                messagebox.showwarning("Error", "Selected vehicle not found in the active records. It may have already been checked in or removed.")
        else:
            messagebox.showwarning("Error", "Vehicle information file not found.")

        self.checkin_win.destroy()
        self.update_table()

    def view_history(self):
        """
        Opens a new Toplevel window to display the vehicle history from 'vehicle_history.csv'.
        Includes a button to clear the history.
        """
        self.history_win = tk.Toplevel(self.root)
        self.history_win.title("Vehicle History")
        self.history_win.geometry("1000x500")

        history_tree_frame = tk.Frame(self.history_win)
        history_tree_frame.pack(side='top', fill='both', expand=True, padx=10, pady=10)

        columns = ["Vehicle", "Purpose", "User", "Checked Out", "Estimated Check In", "Status", "Actual Check In", "Fuel (%)", "Comments", "Mileage at Check In"]
        self.history_tree = ttk.Treeview(history_tree_frame, columns=columns, show='headings')
        for col in columns:
            self.history_tree.heading(col, text=col)
            self.history_tree.column(col, anchor='center', width=110)
        
        vsb = ttk.Scrollbar(history_tree_frame, orient="vertical", command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=vsb.set)
        
        self.history_tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')

        history_btn_frame = tk.Frame(self.history_win)
        history_btn_frame.pack(side='bottom', pady=5)
        
        tk.Button(history_btn_frame, text="Clear History", command=self.confirm_clear_history).pack(padx=5, pady=5)

        self.populate_history_tree()

    def populate_history_tree(self):
        """
        Populates the history Treeview table with data from 'vehicle_history.csv'.
        Clears existing data before reloading.
        """
        for row in self.history_tree.get_children():
            self.history_tree.delete(row)

        history_file_name = "vehicle_history.csv"
        if os.path.exists(history_file_name):
            with open(history_file_name, 'r', newline='') as f:
                reader = csv.reader(f)
                try:
                    next(reader)
                except StopIteration:
                    pass
                for row in reader:
                    while len(row) < len(self.history_tree["columns"]):
                        row.append("")
                    self.history_tree.insert('', tk.END, values=row)
        else:
            messagebox.showinfo("No History", "No history records found.")

    def confirm_clear_history(self):
        """
        Opens a secondary confirmation window to confirm clearing all vehicle history.
        This provides an extra layer of protection against accidental deletion.
        """
        self.confirm_win = tk.Toplevel(self.history_win)
        self.confirm_win.title("Confirm Clear History")
        self.confirm_win.geometry("300x120")
        self.confirm_win.transient(self.history_win)
        self.confirm_win.grab_set()

        tk.Label(self.confirm_win, text="Are you REALLY REALLY SURE you want to clear ALL history?", wraplength=280).pack(pady=10)
        
        button_frame = tk.Frame(self.confirm_win)
        button_frame.pack(pady=5)

        tk.Button(button_frame, text="Yes, I'm sure", command=self.clear_history).pack(side='left', padx=10)
        tk.Button(button_frame, text="No, Cancel", command=self.confirm_win.destroy).pack(side='right', padx=10)

        self.root.wait_window(self.confirm_win)

    def clear_history(self):
        """
        Performs the action of clearing all vehicle history by deleting 'vehicle_history.csv'.
        Called after the user confirms in the 'confirm_clear_history' window.
        """
        history_file_name = "vehicle_history.csv"
        if os.path.exists(history_file_name):
            try:
                os.remove(history_file_name)
                messagebox.showinfo("History Cleared", "Vehicle history has been cleared successfully.")
                self.populate_history_tree()
            except OSError as e:
                messagebox.showerror("Error", f"Failed to clear history: {e}")
        else:
            messagebox.showinfo("No History", "No history file found to clear.")
        
        self.confirm_win.destroy()

    def get_all_vehicles(self):
        """
        Retrieves a list of all unique vehicle IDs from 'vehicle_info.csv'.
        """
        vehicles = set()
        file_name = "vehicle_info.csv"
        if os.path.exists(file_name):
            with open(file_name, 'r', newline='') as file:
                reader = csv.reader(file)
                try:
                    next(reader)
                except StopIteration:
                    return []
                for row in reader:
                    if row:
                        vehicles.add(row[0])
        return sorted(list(vehicles))

    def get_last_checkin_mileage(self, vehicle_id):
        """
        Retrieves the last recorded check-in mileage for a specific vehicle.
        This is taken from the vehicle_info.csv file, which should hold the
        last known mileage for an INACTIVE vehicle.
        """
        file_name = "vehicle_info.csv"
        if os.path.exists(file_name):
            with open(file_name, 'r', newline='') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    if row["Vehicle"] == vehicle_id:
                        mileage_str = row.get("Mileage at Check In", "").strip()
                        if mileage_str:
                            try:
                                return int(mileage_str)
                            except ValueError:
                                return 0
        return 0

    def open_service_tracker_window(self):
        """
        Opens a new Toplevel window for tracking vehicle services.
        Allows adding service items and displays current service status.
        """
        self.service_win = tk.Toplevel(self.root)
        self.service_win.title("Vehicle Service Tracker")
        self.service_win.geometry("1000x700")

        # --- Vehicle Selection Frame ---
        vehicle_selection_frame = tk.LabelFrame(self.service_win, text="Select Vehicle", padx=10, pady=10)
        vehicle_selection_frame.pack(pady=10, padx=10, fill='x')

        tk.Label(vehicle_selection_frame, text="Vehicle:").grid(row=0, column=0, sticky='w', pady=2)
        self.service_vehicle_var_display = tk.StringVar() # This will hold the vehicle selected in the dropdown
        
        self.service_vehicles = self.get_all_vehicles()
        self.service_vehicle_dropdown_display = ttk.Combobox(vehicle_selection_frame, 
                                                            values=self.service_vehicles, 
                                                            textvariable=self.service_vehicle_var_display, 
                                                            state='readonly')
        self.service_vehicle_dropdown_display.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        if self.service_vehicles:
            self.service_vehicle_dropdown_display.set(self.service_vehicles[0])
            # Bind the selection event to update the service tree
            self.service_vehicle_dropdown_display.bind("<<ComboboxSelected>>", self.on_service_vehicle_selected)
        vehicle_selection_frame.grid_columnconfigure(1, weight=1)

        # --- Filter and Sort Controls ---
        filter_sort_frame = tk.LabelFrame(self.service_win, text="Filter & Sort Services", padx=10, pady=10)
        filter_sort_frame.pack(pady=5, padx=10, fill='x')

        tk.Label(filter_sort_frame, text="Status Filter:").grid(row=0, column=0, sticky='w', padx=5)
        self.filter_var = tk.StringVar(value="All")
        self.filter_dropdown = ttk.Combobox(filter_sort_frame, values=["All", "Due", "OK"], textvariable=self.filter_var, state='readonly')
        self.filter_dropdown.grid(row=0, column=1, sticky='ew', padx=5)
        self.filter_dropdown.bind("<<ComboboxSelected>>", self.on_service_filter_sort_change)

        tk.Label(filter_sort_frame, text="Sort By:").grid(row=0, column=2, sticky='w', padx=5)
        self.sort_var = tk.StringVar(value="Service Item")
        # Ensure these match the actual data column names for sorting logic later
        self.sort_dropdown = ttk.Combobox(filter_sort_frame, 
                                            values=["Service Item", "Mileage Interval", "Time Interval (days)", "Last Service Date", "Last Service Mileage", "Status"], 
                                            textvariable=self.sort_var, 
                                            state='readonly')
        self.sort_dropdown.grid(row=0, column=3, sticky='ew', padx=5)
        self.sort_dropdown.bind("<<ComboboxSelected>>", self.on_service_filter_sort_change)

        filter_sort_frame.grid_columnconfigure(1, weight=1)
        filter_sort_frame.grid_columnconfigure(3, weight=1)


        # --- Top Frame for Adding Service Items ---
        add_service_frame = tk.LabelFrame(self.service_win, text="Add/Update Service Item for Selected Vehicle", padx=10, pady=10)
        add_service_frame.pack(pady=10, padx=10, fill='x')

        # The actual variable to be used for saving/updating will come from the dropdown selection
        self.service_item_vehicle_var = tk.StringVar(value=self.service_vehicle_var_display.get())
        self.service_vehicle_var_display.trace_add("write", lambda name, index, mode: self.service_item_vehicle_var.set(self.service_vehicle_var_display.get()))


        tk.Label(add_service_frame, text="Service Item:").grid(row=0, column=0, sticky='w', pady=2)
        self.service_item_entry = tk.Entry(add_service_frame)
        self.service_item_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        self.service_item_entry.bind("<KeyRelease>", self.auto_fill_service_intervals) # Bind for auto-fill

        tk.Label(add_service_frame, text="Mileage Interval (miles):").grid(row=1, column=0, sticky='w', pady=2)
        self.mileage_interval_entry = tk.Entry(add_service_frame)
        self.mileage_interval_entry.grid(row=1, column=1, sticky='ew', padx=5, pady=2)

        tk.Label(add_service_frame, text="Time Interval (days):").grid(row=2, column=0, sticky='w', pady=2)
        self.time_interval_entry = tk.Entry(add_service_frame)
        self.time_interval_entry.grid(row=2, column=1, sticky='ew', padx=5, pady=2)

        tk.Label(add_service_frame, text="Last Service Date (YYYY-MM-DD):").grid(row=3, column=0, sticky='w', pady=2)
        self.last_service_date_entry = tk.Entry(add_service_frame)
        self.last_service_date_entry.grid(row=3, column=1, sticky='ew', padx=5, pady=2)
        self.last_service_date_entry.insert(0, date.today().strftime("%Y-%m-%d"))

        tk.Label(add_service_frame, text="Last Service Mileage:").grid(row=4, column=0, sticky='w', pady=2)
        self.last_service_mileage_entry = tk.Entry(add_service_frame)
        self.last_service_mileage_entry.grid(row=4, column=1, sticky='ew', padx=5, pady=2)

        add_service_button_frame = tk.Frame(add_service_frame)
        add_service_button_frame.grid(row=5, column=0, columnspan=2, pady=10)
        tk.Button(add_service_button_frame, text="Add/Update Service", command=self.add_or_update_service_item).pack(side='left', padx=5)
        tk.Button(add_service_button_frame, text="Mark Service Complete", command=self.mark_service_complete).pack(side='left', padx=5)
        tk.Button(add_service_button_frame, text="Delete Service Config", command=self.delete_service_config).pack(side='left', padx=5)
        
        # --- Service List Table ---
        service_table_frame = tk.Frame(self.service_win)
        service_table_frame.pack(pady=10, padx=10, fill='both', expand=True)

        service_columns = ["Service Item", "Mileage Interval", "Time Interval (days)", "Last Service Date", "Last Service Mileage", "Status"]
        self.service_tree = ttk.Treeview(service_table_frame, columns=service_columns, show='headings', selectmode='browse')
        for col in service_columns:
            self.service_tree.heading(col, text=col)
            self.service_tree.column(col, width=120, anchor='center')
        
        self.service_tree.tag_configure('due_service', background='red', foreground='white')
        self.service_tree.tag_configure('ok_service', background='pale green')

        service_vsb = ttk.Scrollbar(service_table_frame, orient="vertical", command=self.service_tree.yview)
        self.service_tree.configure(yscrollcommand=service_vsb.set)
        
        service_tree_hsb = ttk.Scrollbar(service_table_frame, orient="horizontal", command=self.service_tree.xview)
        self.service_tree.configure(xscrollcommand=service_tree_hsb.set)

        self.service_tree.pack(side='left', fill='both', expand=True)
        service_vsb.pack(side='right', fill='y')
        service_tree_hsb.pack(side='bottom', fill='x')

        self.service_tree.bind("<<TreeviewSelect>>", self.on_service_table_select)

        # Initial population based on the first vehicle or empty
        self.populate_service_tree(self.service_vehicle_var_display.get(), self.filter_var.get(), self.sort_var.get())

    def on_service_vehicle_selected(self, event):
        """
        Event handler for when a vehicle is selected in the service tracker dropdown.
        Updates the service tree to show only that vehicle's service items.
        """
        selected_vehicle = self.service_vehicle_var_display.get()
        # Reset filter and sort when vehicle changes to ensure consistent view
        self.filter_var.set("All") 
        self.sort_var.set("Service Item")
        self.populate_service_tree(selected_vehicle, self.filter_var.get(), self.sort_var.get())
        self.clear_service_entries() # Clear entries when vehicle selection changes

    def on_service_filter_sort_change(self, event):
        """
        Event handler for when filter or sort options are changed.
        Reloads the service tree with the new filtering/sorting.
        """
        selected_vehicle = self.service_vehicle_var_display.get()
        status_filter = self.filter_var.get()
        sort_by = self.sort_var.get()
        self.populate_service_tree(selected_vehicle, status_filter, sort_by)

    def on_service_table_select(self, event):
        """
        Event handler for when a row in the service Treeview table is selected.
        Populates the input fields with the selected service item's data.
        """
        selected = self.service_tree.selection()
        if selected:
            values = self.service_tree.item(selected[0], 'values')
            if values:
                # Vehicle dropdown is already set by on_service_vehicle_selected
                self.service_item_entry.delete(0, tk.END)
                self.mileage_interval_entry.delete(0, tk.END)
                self.time_interval_entry.delete(0, tk.END)
                self.last_service_date_entry.delete(0, tk.END)
                self.last_service_mileage_entry.delete(0, tk.END)

                self.service_item_entry.insert(0, values[0]) # Service Item
                self.mileage_interval_entry.insert(0, values[1]) # Mileage Interval
                self.time_interval_entry.insert(0, values[2]) # Time Interval
                self.last_service_date_entry.insert(0, values[3]) # Last Service Date
                self.last_service_mileage_entry.insert(0, values[4]) # Last Service Mileage

    def auto_fill_service_intervals(self, event=None):
        """
        Auto-fills Mileage Interval and Time Interval if the entered Service Item
        already exists in the vehicle_services.csv, for any vehicle.
        """
        current_service_item = self.service_item_entry.get().strip()
        if not current_service_item:
            return

        file_name = "vehicle_services.csv"
        if os.path.exists(file_name) and os.path.getsize(file_name) > 0:
            with open(file_name, 'r', newline='') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    if row["Service Item"].strip().lower() == current_service_item.lower():
                        self.mileage_interval_entry.delete(0, tk.END)
                        self.mileage_interval_entry.insert(0, row["Mileage Interval (miles)"])
                        
                        self.time_interval_entry.delete(0, tk.END)
                        self.time_interval_entry.insert(0, row["Time Interval (days)"])
                        return # Found a match, stop searching

    def add_or_update_service_item(self):
        """
        Adds a new service item configuration or updates an existing one
        in 'vehicle_services.csv'.
        """
        vehicle = self.service_vehicle_var_display.get().strip() # Get vehicle from the display dropdown
        service_item = self.service_item_entry.get().strip()
        mileage_interval_str = self.mileage_interval_entry.get().strip()
        time_interval_str = self.time_interval_entry.get().strip()
        last_service_date_str = self.last_service_date_entry.get().strip()
        last_service_mileage_str = self.last_service_mileage_entry.get().strip()

        if not vehicle:
            messagebox.showwarning("Input Error", "Please select a vehicle.")
            return

        if not all([service_item, mileage_interval_str, time_interval_str, last_service_date_str, last_service_mileage_str]):
            messagebox.showwarning("Input Error", "Please fill all service item fields.")
            return

        try:
            mileage_interval = int(mileage_interval_str)
            time_interval = int(time_interval_str)
            last_service_mileage = int(last_service_mileage_str)
            if mileage_interval <= 0 or time_interval <= 0 or last_service_mileage < 0:
                raise ValueError("Intervals must be positive, mileage non-negative.")
        except ValueError:
            messagebox.showwarning("Input Error", "Mileage/Time Intervals and Last Service Mileage must be valid numbers.")
            return
        
        last_service_date = self.try_parse_date(last_service_date_str)
        if not last_service_date:
            messagebox.showwarning("Date Error", "Invalid Last Service Date format. Please use YYYY-MM-DD.")
            return

        file_name = "vehicle_services.csv"
        data = []
        if os.path.exists(file_name) and os.path.getsize(file_name) > 0:
            with open(file_name, 'r', newline='') as f:
                reader = csv.reader(f)
                data = list(reader)
        
        header = ["Vehicle", "Service Item", "Mileage Interval (miles)", "Time Interval (days)", "Last Service Date (YYYY-MM-DD)", "Last Service Mileage"]
        if not data or data[0] != header:
            data = [header]

        new_row = [vehicle, service_item, str(mileage_interval), str(time_interval), 
                   last_service_date.strftime("%Y-%m-%d"), str(last_service_mileage)]

        found = False
        for i in range(1, len(data)):
            if data[i][0] == vehicle and data[i][1] == service_item:
                data[i] = new_row
                found = True
                break
        
        if not found:
            data.append(new_row)

        with open(file_name, 'w', newline='') as f:
            csv.writer(f).writerows(data)
        
        messagebox.showinfo("Success", "Service item added/updated successfully!")
        self.clear_service_entries()
        self.populate_service_tree(vehicle, self.filter_var.get(), self.sort_var.get()) # Refresh for the current vehicle

    def mark_service_complete(self):
        """
        Marks the selected service item as complete by updating its last service date and mileage.
        """
        selected = self.service_tree.selection()
        if not selected:
            messagebox.showwarning("Selection Error", "No service item selected to mark complete.")
            return

        values = self.service_tree.item(selected[0], 'values')
        vehicle = self.service_vehicle_var_display.get() # Get vehicle from the display dropdown
        service_item = values[0] # Service Item is the first value in the filtered tree

        current_mileage_str = simpledialog.askstring("Mark Service Complete", f"Enter current mileage for {vehicle} - {service_item}:", parent=self.service_win)
        if current_mileage_str is None:
            return
        
        try:
            current_mileage = int(current_mileage_str.strip())
            if current_mileage < 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Input Error", "Invalid mileage entered. Please enter a valid non-negative number.")
            return
        
        today_date_str = date.today().strftime("%Y-%m-%d")

        file_name = "vehicle_services.csv"
        data = []
        if os.path.exists(file_name):
            with open(file_name, 'r', newline='') as f:
                reader = csv.reader(f)
                data = list(reader)
        
        found = False
        for i in range(1, len(data)):
            if data[i][0] == vehicle and data[i][1] == service_item:
                data[i][4] = today_date_str
                data[i][5] = str(current_mileage)
                found = True
                break
        
        if found:
            with open(file_name, 'w', newline='') as f:
                csv.writer(f).writerows(data)
            messagebox.showinfo("Success", f"Service '{service_item}' for '{vehicle}' marked complete.")
            self.populate_service_tree(vehicle, self.filter_var.get(), self.sort_var.get()) # Refresh for the current vehicle
        else:
            messagebox.showwarning("Error", "Selected service item not found in records.")

    def delete_service_config(self):
        """
        Deletes the selected service item configuration from 'vehicle_services.csv'.
        """
        selected = self.service_tree.selection()
        if not selected:
            messagebox.showwarning("Selection Error", "No service item selected to delete.")
            return

        values = self.service_tree.item(selected[0], 'values')
        vehicle = self.service_vehicle_var_display.get() # Get vehicle from the display dropdown
        service_item = values[0] # Service Item is the first value in the filtered tree

        confirm = messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete the service configuration for '{service_item}' on '{vehicle}'?")
        if not confirm:
            return

        file_name = "vehicle_services.csv"
        data = []
        if os.path.exists(file_name):
            with open(file_name, 'r', newline='') as f:
                reader = csv.reader(f)
                data = list(reader)
        
        new_data = [data[0]]
        found = False
        for i in range(1, len(data)):
            if not (data[i][0] == vehicle and data[i][1] == service_item):
                new_data.append(data[i])
            else:
                found = True
        
        if found:
            with open(file_name, 'w', newline='') as f:
                csv.writer(f).writerows(new_data)
            messagebox.showinfo("Success", f"Service configuration for '{service_item}' on '{vehicle}' deleted.")
            self.populate_service_tree(vehicle, self.filter_var.get(), self.sort_var.get()) # Refresh for the current vehicle
        else:
            messagebox.showwarning("Error", "Selected service configuration not found.")


    def clear_service_entries(self):
        """Clears the text in the service input Entry widgets."""
        self.service_item_entry.delete(0, tk.END)
        self.mileage_interval_entry.delete(0, tk.END)
        self.time_interval_entry.delete(0, tk.END)
        self.last_service_date_entry.delete(0, tk.END)
        self.last_service_mileage_entry.delete(0, tk.END)
        self.last_service_date_entry.insert(0, date.today().strftime("%Y-%m-%d"))

    def populate_service_tree(self, vehicle_filter=None, status_filter="All", sort_by="Service Item"):
        """
        Populates the vehicle service Treeview table with data from 'vehicle_services.csv'.
        Applies styling based on whether a service is due.
        Filters by vehicle_filter, status_filter and sorts by sort_by.
        """
        for row in self.service_tree.get_children():
            self.service_tree.delete(row)

        service_file_name = "vehicle_services.csv"
        
        # Ensure header is always present if file is empty but exists
        self.init_service_file() 

        all_service_data = []
        if os.path.exists(service_file_name) and os.path.getsize(service_file_name) > 0:
            with open(service_file_name, 'r', newline='') as f:
                reader = csv.DictReader(f)
                all_service_data = list(reader)
            
            # 1. Filter by selected vehicle
            filtered_by_vehicle = [row for row in all_service_data if row["Vehicle"] == vehicle_filter]

            if not filtered_by_vehicle and vehicle_filter:
                self.service_tree.insert('', tk.END, values=["No service items configured for this vehicle."])
                return

            # 2. Calculate status for each item and apply status filter
            processed_data = []
            for row_data in filtered_by_vehicle:
                vehicle = row_data["Vehicle"] 
                service_item = row_data["Service Item"]
                mileage_interval = int(row_data["Mileage Interval (miles)"])
                time_interval = int(row_data["Time Interval (days)"])
                last_service_date_str = row_data["Last Service Date (YYYY-MM-DD)"]
                last_service_mileage = int(row_data["Last Service Mileage"])

                last_service_date = self.try_parse_date(last_service_date_str)
                
                current_vehicle_mileage = self.get_last_checkin_mileage(vehicle)
                
                is_due = False
                status_text = "OK"

                if last_service_date:
                    due_date = last_service_date + timedelta(days=time_interval)
                    if date.today() > due_date:
                        is_due = True
                        status_text = "DUE (Time)"
                else:
                    is_due = True
                    status_text = "DUE (Date Missing)" 

                if not is_due and current_vehicle_mileage is not None: 
                    if (current_vehicle_mileage - last_service_mileage) >= mileage_interval:
                        is_due = True
                        status_text = "DUE (Mileage)"
                
                if (last_service_date and date.today() > (last_service_date + timedelta(days=time_interval))) and \
                   (current_vehicle_mileage is not None and (current_vehicle_mileage - last_service_mileage) >= mileage_interval):
                    status_text = "DUE (Time & Mileage)"

                # Add status to the row data for filtering/sorting purposes
                row_data['Calculated Status'] = status_text
                row_data['Is Due'] = is_due
                processed_data.append(row_data)

            final_filtered_data = []
            for item in processed_data:
                if status_filter == "All":
                    final_filtered_data.append(item)
                elif status_filter == "Due" and item['Is Due']:
                    final_filtered_data.append(item)
                elif status_filter == "OK" and not item['Is Due']:
                    final_filtered_data.append(item)

            # 3. Sort the data
            def get_sort_key(item):
                if sort_by == "Service Item":
                    return item["Service Item"].lower()
                elif sort_by == "Mileage Interval":
                    return int(item["Mileage Interval (miles)"])
                elif sort_by == "Time Interval (days)":
                    return int(item["Time Interval (days)"])
                elif sort_by == "Last Service Date":
                    return self.try_parse_date(item["Last Service Date (YYYY-MM-DD)"]) or date.min
                elif sort_by == "Last Service Mileage":
                    return int(item["Last Service Mileage"])
                elif sort_by == "Status":
                    # Sort Due items before OK items, then alphabetically for consistency
                    if item['Is Due']:
                        return (0, item['Calculated Status']) 
                    else:
                        return (1, item['Calculated Status'])
                return item["Service Item"].lower() # Default sort

            final_filtered_data.sort(key=get_sort_key)


            # 4. Insert into Treeview
            for row_data in final_filtered_data:
                values = [row_data["Service Item"], row_data["Mileage Interval (miles)"], row_data["Time Interval (days)"], 
                          row_data["Last Service Date (YYYY-MM-DD)"], row_data["Last Service Mileage"], row_data['Calculated Status']]
                
                tag = 'due_service' if row_data['Is Due'] else 'ok_service'
                self.service_tree.insert('', tk.END, values=values, tags=(tag,))
        else:
            if vehicle_filter:
                self.service_tree.insert('', tk.END, values=[f"No service records found for '{vehicle_filter}'. Add a service item above."])
            else:
                self.service_tree.insert('', tk.END, values=["No vehicle service records found. Select a vehicle or add one first."])


if __name__ == "__main__":
    root = tk.Tk()
    app = VehicleManager(root)
    root.mainloop()