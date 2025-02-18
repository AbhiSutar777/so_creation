// Copyright (c) 2025, Abhijeet Sutar and contributors
// For license information, please see license.txt

// frappe.ui.form.on("Upload Data", {
// 	refresh(frm) {

// 	},
// });

GET_UPLOAD_TEMPLATE = "so_creation.so_creation.doctype.upload_data.upload_data.download_excel_file"
CREATE_SALES_ORDER = "so_creation.so_creation.doctype.upload_data.upload_data.create_sales_order_from_excel"
// GET_BRANCH_LIST = "nelson.nelson.doctype.manual_sales_forecast_update.manual_sales_forecast_update.get_customer_specific_branch"

frappe.ui.form.on("Upload Data", {

	download_template(frm){
		frappe.call({
			method: GET_UPLOAD_TEMPLATE,
			args: {'data': frm.doc},
			callback: function(r) {
				frappe.show_alert({
				    message:__('Downloading template'),
				    indicator:'green'
				}, 3);
				window.location.href = r.message.file_url;
			}
		});
		frm.set_value("attach_file", "")
	},


	upload_data: function(frm) {
	    if (!frm.doc.attach_file) {
	        frappe.throw("Please attach file first");
	    }
	    frappe.confirm('This process will erase all the previous uploaded data in Annexure I. Are you sure you want to proceed ahead?', () => {
	        frappe.show_progress("Processing", 0, 100, "Starting...");
	        frappe.call({
	            method: CREATE_SALES_ORDER,
	            args: { 'doc': frm.doc },
	            callback: function(r) {
	                frappe.show_alert({
	                    message: __('Creating Sales Forecast Upload Records'),
	                    indicator: 'green'
	                }, 3);
	                window.location.href = r.message.file_url;
	            }
	        });
	
	        frappe.realtime.on('progress_update', function(data) {
	            frappe.show_progress("Processing", data.progress, data.total, data.message);
	        });
	    }, () => {});
	},

});


