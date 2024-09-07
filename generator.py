import tkinter as tk
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime
from tkinter import messagebox

# Function to clear item fields after adding an item
def clear_item():
    qty_spinbox.delete(0, tk.END)
    qty_spinbox.insert(0, "1")
    desc_entry.delete(0, tk.END)
    price_spinbox.delete(0, tk.END)
    price_spinbox.insert(0, "0.0")

# Function to reset all fields
def reset_fields():
    """Reset all fields and clear the invoice list"""
    first_name_entry.delete(0, tk.END)
    last_name_entry.delete(0, tk.END)
    phone_entry.delete(0, tk.END)
    discount_entry.delete(0, tk.END)
    clear_item()
    tree.delete(*tree.get_children())
    invoice_list.clear()

# Function to add item to the treeview and invoice list
def add_item():
    try:
        qty = int(qty_spinbox.get())
        desc = desc_entry.get().strip()
        price = float(price_spinbox.get())

        if not desc:
            messagebox.showwarning("Input Error", "Description cannot be empty!")
            return

        line_total = qty * price
        invoice_item = [qty, desc, price, line_total]

        tree.insert('', 0, values=invoice_item)
        clear_item()
        invoice_list.append(invoice_item)

    except ValueError:
        messagebox.showerror("Input Error", "Please enter valid numeric values for Qty and Unit Price.")

# Function to generate the invoice with a discount
def generate_invoice():
    if not first_name_entry.get().strip() or not last_name_entry.get().strip() or not phone_entry.get().strip():
        messagebox.showwarning("Input Error", "Please provide customer details before generating the invoice.")
        return

    if not invoice_list:
        messagebox.showwarning("Input Error", "No items in the invoice.")
        return

    # Gather customer info
    doc = DocxTemplate("invoice_template.docx")
    name = f"{first_name_entry.get().strip()} {last_name_entry.get().strip()}"
    phone = phone_entry.get().strip()

    # Calculate subtotal
    subtotal = sum(item[3] for item in invoice_list)
    
    # Handle Discount
    try:
        discount = float(discount_entry.get().strip()) if discount_entry.get().strip() else 0
    except ValueError:
        messagebox.showerror("Input Error", "Invalid discount value.")
        return

    if discount < 0 or discount > subtotal:
        messagebox.showerror("Input Error", "Discount cannot be greater than the subtotal or negative.")
        return

    # Calculate tax and total after discount
    salestax = 0.1
    total_after_discount = subtotal - discount
    total = total_after_discount * (1 + salestax)

    # Render the invoice
    doc.render({
        "name": name,
        "phone": phone,
        "invoice_list": invoice_list,
        "subtotal": subtotal,
        "discount": discount,
        "total_after_discount": total_after_discount,
        "salestax": f"{salestax * 100}%",
        "total": total
    })

    # Save the document with a timestamp
    doc_name = f"invoice_{name}_{datetime.datetime.now().strftime('%Y-%m-%d_%H%M%S')}.docx"
    doc.save(doc_name)

    messagebox.showinfo("Invoice Complete", f"Invoice '{doc_name}' generated successfully!")
    reset_fields()

# Initialize the tkinter window
window = tk.Tk()
window.title("Invoice Generator with Discount")

# Create the main frame
frame = tk.Frame(window)
frame.pack(padx=20, pady=10)

# Customer Info
tk.Label(frame, text="First Name").grid(row=0, column=0)
first_name_entry = tk.Entry(frame)
first_name_entry.grid(row=1, column=0)

tk.Label(frame, text="Last Name").grid(row=0, column=1)
last_name_entry = tk.Entry(frame)
last_name_entry.grid(row=1, column=1)

tk.Label(frame, text="Phone").grid(row=0, column=2)
phone_entry = tk.Entry(frame)
phone_entry.grid(row=1, column=2)

# Item Info
tk.Label(frame, text="Qty").grid(row=2, column=0)
qty_spinbox = tk.Spinbox(frame, from_=1, to=100)
qty_spinbox.grid(row=3, column=0)

tk.Label(frame, text="Description").grid(row=2, column=1)
desc_entry = tk.Entry(frame)
desc_entry.grid(row=3, column=1)

tk.Label(frame, text="Unit Price").grid(row=2, column=2)
price_spinbox = tk.Spinbox(frame, from_=0.0, to=500, increment=0.5)
price_spinbox.grid(row=3, column=2)

# Discount Field
tk.Label(frame, text="Discount").grid(row=4, column=0)
discount_entry = tk.Entry(frame)
discount_entry.grid(row=5, column=0)

# Add Item Button
tk.Button(frame, text="Add Item", command=add_item).grid(row=5, column=2, pady=5)

# Invoice Table
columns = ('qty', 'desc', 'price', 'total')
tree = ttk.Treeview(frame, columns=columns, show="headings")
for col in columns:
    tree.heading(col, text=col.title())
tree.grid(row=6, column=0, columnspan=3, padx=20, pady=10)

# Generate and New Invoice Buttons
tk.Button(frame, text="Generate Invoice", command=generate_invoice).grid(row=7, column=0, columnspan=3, sticky="nsew", padx=20, pady=5)
tk.Button(frame, text="New Invoice", command=reset_fields).grid(row=8, column=0, columnspan=3, sticky="nsew", padx=20, pady=5)

invoice_list = []
window.mainloop()
