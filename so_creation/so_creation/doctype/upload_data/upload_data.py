# Copyright (c) 2025, Abhijeet Sutar and contributors
# For license information, please see license.txt

# import frappe
from frappe.model.document import Document


class UploadData(Document):
	pass

import os
import json
import hashlib
import csv
import frappe
import pandas as pd
import openpyxl
from datetime import datetime, timedelta
from frappe import _
from frappe.utils import get_files_path, get_url, nowdate, get_site_path
from frappe.model.document import Document
from frappe.utils.file_manager import save_file
from frappe.handler import download_file
from openpyxl import Workbook
from openpyxl.styles import numbers


# from nelson.nelson.utils import read_data_from_excel


class ManualSalesForecastUpdate(Document):
	pass

UPLOAD_TEMPLATE_COLUMNS = ["purch_doc","product_id","quantity","uom","net_price","crcy","per","corporate_name","doc_date","customer_name"]

def read_data_from_excel(file_path, field_list):
    try:
        # Read the Excel file
        df = pd.read_excel(file_path, header=0, usecols=field_list)
        df.fillna('', inplace=True)
        # Convert to list
        data_list = []
        for index, row in df.iterrows():
            row_dict = {}
            for col_name in field_list:
                row_dict[col_name] = row[col_name]
            data_list.append(row_dict)
        return data_list

    except Exception as e:
        raise e

def create_manual_sales_forecast_update(data):
    # records = get_party_specific_items(data)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "data"

    # Hardcoded headers if something is changed update here
    sheet.append(["purch_doc","product_id","quantity","uom","net_price","crcy","per","corporate_name","doc_date","customer_name"])

    name = f"upload_data_template"
    file_path = f"/tmp/{name}_.xlsx"
    
    # Save the workbook to the file path
    workbook.save(file_path)
    return file_path

##-----------------------------------------------------------------------------------------------------------------------

@frappe.whitelist()
def download_excel_file(data):
    data = json.loads(data)
    file_path = create_manual_sales_forecast_update(data)

    with open(file_path, "rb") as filedata:
        file_name = os.path.basename(file_path)  # Extract the file name from the path
        content = filedata.read()

    # Check if the file already exists in the files directory
    file_path_in_files = os.path.join(get_files_path(), file_name)

    # If file exists, rename it by appending a suffix
    if os.path.exists(file_path_in_files):
        base_name, extension = os.path.splitext(file_name)
        i = 1
        while os.path.exists(os.path.join(get_files_path(), f"{base_name}_{i}{extension}")):
            i += 1
        file_name = f"{base_name}_{i}{extension}"

    # Save file in the File Manager with the adjusted name
    file_object = save_file(file_name, content, "Upload Data", 'Upload Data', is_private=False)

    file_url = get_url() + "/files/" + file_object.file_name
    return {"file_url": file_url}


##----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

@frappe.whitelist()
def set_doc_name(doc, method=None):
	if doc.is_new():
		# Read excel data
		file_path = get_url() + doc.attach_file
		records = read_data_from_excel(file_path, UPLOAD_TEMPLATE_COLUMNS)
		##--------------------------------------------------------------------------------------------------------------------
		## Get PO NUMBER...
		po_no_list = []
		for r in records:
			if r['purch_doc'] not in po_no_list:
				po_no_list.append(r['purch_doc'])
				po_no = ''
		if len(po_no_list) > 1:
			frappe.throw(f"You cannot include more than 1 customer in one Upload Templante.")
		else:
			po_no = po_no_list[0]
			doc.name = po_no



##----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
@frappe.whitelist()
def process_file(data):
    # data = json.loads(data)
    
    # Read excel data
    file_path = get_url() + data.get("attach_file")
    records = read_data_from_excel(file_path, UPLOAD_TEMPLATE_COLUMNS)
    ##--------------------------------------------------------------------------------------------------------------------
    ## Get PO NUMBER...
    po_no_list = []
    for r in records:
    	if r['purch_doc'] not in po_no_list:
    		po_no_list.append(r['purch_doc'])

    po_no = ''
    if len(po_no_list) > 1:
    	frappe.throw(f"You cannot include more than 1 customer in one Upload Templante.")
    else:
    	po_no = po_no_list[0]
    
    # frappe.msgprint(f"PO Number : {po_no}")

    ##--------------------------------------------------------------------------------------------------------------------
    ## Get CUSTOMER NAME...
    customer_list = []
    for r in records:
    	if r['customer_name'] not in customer_list:
    		customer_list.append(r['customer_name'])

    customer_name = ''
    if len(customer_list) > 1:
    	frappe.throw(f"You cannot include more than 1 customer in one Upload Templante.")
    else:
    	customer_name = customer_list[0]
    
    # frappe.msgprint(f"Customer Name : {customer_name}")

    ##--------------------------------------------------------------------------------------------------------------------
	## Get CUSTOMER NAME...
    po_date_list = []
    for r in records:
    	if r['doc_date'] not in po_date_list:
    		po_date_list.append(r['doc_date'])

    po_date = ''
    if len(po_date_list) > 1:
    	frappe.throw(f"You cannot include more than 1 customer in one Upload Templante.")
    else:
    	po_date = po_date_list[0]
    
    # frappe.msgprint(f"Customer Name : {po_date}")

    ##--------------------------------------------------------------------------------------------------------------------
	## Get ITEM TABLE...
    item_table = []
    for r in records:
    	# found = False
    	# for i in item_table:
    	# 	if r['product_id'] == i['item_code'] and r['quantity'] == i['qty']:
    	# 		found = True
    	# 	else:
    	# 		pass
    	# if found == False:
    	item_table.append({
            "item_name": r["product_id"],
            "qty": r["quantity"],
            "rate": r["net_price"],
            "uom": r["uom"]
        })
    # frappe.msgprint(f"Item Table : {item_table}")

    return customer_name, po_no, po_date, item_table

###========================================================================================================================
###========================================================================================================================
###========================================================================================================================

@frappe.whitelist()
def create_sales_order_from_excel(doc, method=None):
    # # Parse the incoming document if it's a string
    doc = frappe.parse_json(doc)

    # # Load the actual Document object using its doctype and name
    doc = frappe.get_doc(doc.get('doctype'), doc.get('name'))

    # # Call the process_po_data function to extract the data
    customer_name, po_no, po_date, item_table = process_file(doc)

    # # Now call the create_sales_order with the extracted values
    create_sales_order(customer_name, po_no, po_date, item_table)


## Create SO
def create_sales_order(customer_name, po_no, po_date, item_table):
    """
    Function to create a Sales Order from the extracted PO data.
    """

    # # Convert required_date to a date object for comparison
    required_date = frappe.utils.getdate(frappe.utils.today())
    today = frappe.utils.getdate(frappe.utils.today())

    # # frappe.msgprint(f"Converted Required Date: {required_date}")
    
    # # Compare the required_date with today's date
    if required_date < today:
        customer_requirement_date = today
    else:
        customer_requirement_date = required_date

    # # Check if customer exists, if not create new
    customer = get_or_create_customer(customer_name)
    # customer = customer_name

    # # Prepare item table for sales order
    items_list = []
    for idx, item in enumerate(item_table):
    # # for item in item_table:
        item_name = item.get("item_name")
        qty = float(item.get("qty", 0))  # Default to 0 if quantity is missing
        rate = item.get("rate", "0")  # Default to "0" if rate is missing
        uom = item.get("uom", "Nos")

        # # Check if item exists, if not create new
        item_code = get_or_create_item(item_name)

        # # Ensure UOM exists, if not create it
        uom_name = get_or_create_uom(uom)  # Pass the cleaned UOM

        # # Add the item to the sales order item list
        items_list.append({
            "item_code": item_code,
            "qty": qty,
            "rate": rate,  # Convert cleaned rate to float
            "uom": uom_name,  # Include UOM in the Sales Order Item
            "delivery_date" : customer_requirement_date
        })

    # # Create a new Sales Order
    sales_order = frappe.get_doc({
        "doctype": "Sales Order",
        "customer": customer,
        "po_no": po_no,
        "po_date": frappe.utils.getdate(po_date),
        "delivery_date": customer_requirement_date,
        "items": items_list,
    })

    # # Save and optionally submit the Sales Order
    sales_order.save()
    # sales_order.submit()  # Uncomment to submit the Sales Order

    # # Show success message
    frappe.msgprint(
        f"Sales Order has been created successfully. Click on "
        f"<a href='http://techno2.dev.vedarthsolutions.com/app/sales-order/{sales_order.name}' target='_blank'>{sales_order.name}</a> to view the Sales Order."
    )
    return sales_order.name

###-----------------------------------------------------------------------------------------------------------------------------------------------
###-----------------------------------------------------------------------------------------------------------------------------------------------

# # # Check If UOM exists in system or not------------------------------------------------------------------------------------------------------------------------------
def get_or_create_uom(uom_name):
    """
    Check if a Unit of Measure (UOM) exists by name, if not, create a new UOM.
    """
    # # Standardize the UOM name to uppercase for consistent checks
    uom_name = uom_name.upper()

    # # Check if the UOM exists in a case-insensitive way
    if not frappe.db.exists("UOM", {"uom_name": uom_name}):
        uom = frappe.new_doc("UOM")
        uom.uom_name = uom_name
        uom.save()
        frappe.msgprint(f"New UOM '{uom_name}' created.")

    return uom_name  # Return the standardized UOM name

# # # Check If Customer exists in system or not------------------------------------------------------------------------------------------------------------------------------
def get_or_create_customer(customer_name):
    """
    Check if a customer exists by name, if not, create a new customer.
    """
    if frappe.db.exists("Customer", {"customer_name": customer_name}):
        customer = frappe.get_doc("Customer", {"customer_name": customer_name})
    else:
        customer = frappe.new_doc("Customer")
        customer.customer_name = customer_name
        customer.save()
        frappe.msgprint(f"New customer '{customer_name}' created.")

    return customer

# # # Check If Item exists in system or not------------------------------------------------------------------------------------------------------------------------------
def get_or_create_item(item_name):
    """
    Check if an item exists by name, if not, create a new item.
    """
    if frappe.db.exists("Item", {"item_name": ['like', f"%{item_name}%"]}):
        item = frappe.get_doc("Item", {"item_name": ['like', f"%{item_name}%"]})
    else:
        item = frappe.new_doc("Item")
        item.item_name = item_name
        item.item_code = item_name
        item.item_group = "All Item Groups"  # Set the default item group
        item.save()
        frappe.msgprint(f"New item '{item_name}' created.")

    return item.item_code  # Return the item_code