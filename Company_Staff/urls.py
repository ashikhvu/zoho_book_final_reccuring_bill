#zoho Final
from django.urls import path,re_path
from . import views
from django.conf.urls.static import static
from django.conf import settings
from django.contrib.staticfiles.urls import staticfiles_urlpatterns
from django.views.static import serve

urlpatterns = [
    # -------------------------------Company section--------------------------------
    path('Company/Dashboard',views.company_dashboard,name='company_dashboard'),
    path('Company/Staff-Request',views.company_staff_request,name='company_staff_request'),
    path('Company/Staff-Request/Accept/<int:pk>',views.staff_request_accept,name='staff_request_accept'),
    path('Company/Staff-Request/Reject/<int:pk>',views.staff_request_reject,name='staff_request_reject'),
    path('Company/All-Staffs',views.company_all_staff,name='company_all_staff'),
    path('Company/Staff-Approval/Cancel/<int:pk>',views.staff_approval_cancel,name='staff_approval_cancel'),
    path('Company/Profile',views.company_profile,name='company_profile'),
    path('Company/Profile-Editpage',views.company_profile_editpage,name='company_profile_editpage'),
    path('Company/Profile/Edit/Basicdetails',views.company_profile_basicdetails_edit,name='company_profile_basicdetails_edit'),
    path('Company/Password_Change',views.company_password_change,name='company_password_change'),
    path('Company/Profile/Edit/Companydetails',views.company_profile_companydetails_edit,name='company_profile_companydetails_edit'),
    path('Company/Module-Editpage',views.company_module_editpage,name='company_module_editpage'),
    path('Company/Module-Edit',views.company_module_edit,name='company_module_edit'),
    path('Company/Renew/Payment_terms',views.company_renew_terms,name='company_renew_terms'),
    path('Company/Notifications',views.company_notifications,name='company_notifications'),
    path('company/messages/read/<int:pk>',views.company_message_read,name='company_message_read'),
    path('Company/Payment_History',views.company_payment_history,name='company_payment_history'),
    path('Company/Trial/Review',views.company_trial_feedback,name='company_trial_feedback'),
    path('Company/Profile/Edit/gsttype',views.company_gsttype_change,name='company_gsttype_change'),


    # -------------------------------Staff section--------------------------------
    path('Staff/Dashboard',views.staff_dashboard,name='staff_dashboard'),
    path('Staff/Profile',views.staff_profile,name='staff_profile'),
    path('Staff/Profile-Editpage',views.staff_profile_editpage,name='staff_profile_editpage'),
    path('Staff/Profile/Edit/details',views.staff_profile_details_edit,name='staff_profile_details_edit'),
    path('Staff/Password_Change',views.staff_password_change,name='staff_password_change'),
    # -------------------------------Zoho Modules section--------------------------------
    
    # ------------------------- TINTO urls items  START---------------------

    path('new_items',views.new_items,name='new_items'),
    path('items_list',views.items_list,name='items_list'),
    path('create_item',views.create_item,name='create_item'),
    path('itemsoverview/<int:pk>',views.itemsoverview,name='itemsoverview'),
    path('edititems/<int:pr>',views.edititems,name='edititems'),
    path('item_status_edit/<int:pv>',views.item_status_edit,name='item_status_edit'),
    path('shareItemToEmail/<int:pt>',views.shareItemToEmail,name='shareItemToEmail'),
    path('deleteitem/<int:pl>',views.deleteitem,name='deleteitem'),
    path('add_item_comment/<int:pc>',views.add_item_comment,name='add_item_comment'),
    path('delete_item_comment/<int:ph>/<int:pr>',views.delete_item_comment,name='delete_item_comment'),
    path('add_unit',views.add_unit,name='add_unit'),
    path('unit_dropdown',views.unit_dropdown,name = 'unit_dropdown'),
    path('downloadItemSampleImportFile',views.downloadItemSampleImportFile,name='downloadItemSampleImportFile'),
    path('import_item',views.import_item,name='import_item'),
    path('item_view_sort_by_name/<int:pk>',views.item_view_sort_by_name,name='item_view_sort_by_name'),
    path('item_view_sort_by_hsn/<int:pk>',views.item_view_sort_by_hsn,name='item_view_sort_by_hsn'),
    path('filter_item_view_Active/<int:pk>',views.filter_item_view_Active,name='filter_item_view_Active'),
    path('filter_item_view_inActive/<int:pk>',views.filter_item_view_inActive,name='filter_item_view_inActive'),

    #----------------------------- TINTO urls items  END-----------------------------

    #-------------------------TINTO Chartof accounts urls  START------------------------

    path('chartofaccounts',views.chartofaccounts,name='chartofaccounts'),
    path('addchartofaccounts',views.addchartofaccounts,name='addchartofaccounts'),
    path('create_account',views.create_account,name='create_account'),
    path('chartofaccountsoverview/<int:pk>',views.chartofaccountsoverview,name='chartofaccountsoverview'),
    path('editchartofaccounts/<int:pr>',views.editchartofaccounts,name='editchartofaccounts'),
    path('deleteaccount/<int:pl>',views.deleteaccount,name='deleteaccount'),
    path('acc_status_edit/<int:pv>',views.acc_status_edit,name='acc_status_edit'),
    path('add_account_comment/<int:pc>',views.add_account_comment,name='add_account_comment'),

    path('delete_account_comment/<int:ph>/<int:pr>',views.delete_account_comment,name='delete_account_comment'),
    path('add_account',views.add_account,name='add_account'),
    path('account_dropdown',views.account_dropdown,name = 'account_dropdown'),
    path('account_view_sort_by_name/<int:pk>',views.account_view_sort_by_name,name='account_view_sort_by_name'),
    path('shareaccountToEmail/<int:pt>',views.shareaccountToEmail,name='shareaccountToEmail'),

    #------------------------- TINTO Chartof accounts urls  ENDS----------------------
    
    path('chartofaccountsActive',views.chartofaccountsActive,name='chartofaccountsActive'),
    path('chartofaccountsInactive',views.chartofaccountsInactive,name='chartofaccountsInactive'),
    
    #---------------------------------Payroll employee-----------------------------------
   #--------------------------------------George Mathew---------------------------------
    path('Company/payroll_employee_create',views.payroll_employee_create,name='payroll_employee_create'),
    path('Company/payroll_employee_list',views.employee_list,name='employee_list'),
    path('Company/payroll_employee_overview/<int:pk>',views.employee_overview,name='employee_overview'),
    path('Company/create_employee',views.create_employee,name='create_employee'),
    path('Company/payroll_employee_edit/<int:pk>',views.payroll_employee_edit,name='payroll_employee_edit'),
    path('Company/do_payroll_edit/<int:pk>',views.do_payroll_edit,name='do_payroll_edit'),
    path('Company/add_comment/<int:pk>',views.add_comment,name='add_comment'),
    path('Company/delete_comment/<int:pk>/<int:pi>',views.delete_commet,name='delete_comment'),
    path('Company/delete_employee/<int:pk>',views.delete_employee,name='delete_employee'),
    path('Company/employee_status/<int:pk>',views.employee_status,name='employee_status'),
    path('Company/add_blood',views.add_blood,name='add_blood'),
    path('company/import_payroll_excel',views.import_payroll_excel,name='import_payroll_excel'),
    path('Company/add_file/<int:pk>',views.add_file,name='add_file'),
    path('company/shareemail/<int:pk>',views.shareemail,name='shareemail'),
#----------------------------------------------------end--------------------------------------------------

    path('accounts_asset_filter',views.accounts_asset_filter,name='accounts_asset_filter'),
    path('accounts_liability_filter',views.accounts_liability_filter,name='accounts_liability_filter'),
    path('accounts_equity_filter',views.accounts_equity_filter,name='accounts_equity_filter'),
    path('accounts_income_filter',views.accounts_income_filter,name='accounts_income_filter'),
    path('accounts_expense_filter',views.accounts_expense_filter,name='accounts_expense_filter'),
    
    path('account_view_sort_by_namelist',views.account_view_sort_by_namelist,name='account_view_sort_by_namelist'),
    
    path('account_view_filterActive/<int:ph>',views.account_view_filterActive,name='account_view_filterActive'),
    path('account_view_filterinActive/<int:ph>',views.account_view_filterinActive,name='account_view_filterinActive'),
    
    #---------------- Banking ------------------#
    path('Company/Banking/List',views.bank_list, name='bank_list'),
    path('Company/Banking/Create',views.load_bank_create, name='load_bank_create'),
    path('Company/Banking/Create/Bank',views.bank_create, name='bank_create'),
    path('Company/Banking/Edit/<int:id>',views.bank_edit, name='bank_edit'),
    path('Company/Banking/Edit/Bank/<int:id>',views.load_bank_edit, name='load_bank_edit'),
    path('Company/Banking/View/Bank/<int:id>',views.bank_view, name='bank_view'),
    path('Company/Banking/Bank/Status/<int:id>',views.banking_status, name='banking_status'),
    path('Company/Banking/Bank/File/<int:id>',views.bank_attachfile, name='bank_attachfile'),
    path('Company/Banking/Delete/Bank/<int:id>',views.delete_banking, name='delete_banking'),
    path('Company/Banking/Send/Bank/<int:id>',views.send_bank_transaction, name='send_bank_transaction'),
    path('Company/Banking/Create/Tranaction/<int:id>',views.bank_transaction_create, name='bank_transaction_create'),
    path('Company/Banking/Details/Tranaction',views.load_trans_details, name='load_trans_details'),
    path('Company/Banking/Edit/Tranaction',views.bank_transaction_edit, name='bank_transaction_edit'),
    path('Company/Banking/Delete/Tranaction/<int:id>',views.delete_transaction, name='delete_transaction'),
    path('Company/Banking/History/<int:id>',views.load_bank_history, name='load_bank_history'),
    path('Company/Banking/Transaction/History/<int:id>',views.load_bank_trans_history, name='load_bank_trans_history'),
    
    #----------------------------------------------------------akshay--start--------------------------------------------------------
    #------------price lists-------------------
    path('all_price_lists', views.all_price_lists, name='all_price_lists'),
    path('create_price_list/', views.create_price_list, name='create_price_list'),
    path('price_list_details/<int:price_list_id>/', views.price_list_details, name='price_list_details'),
    path('edit_price_list/<int:price_list_id>/', views.edit_price_list, name='edit_price_list'),
    path('delete_price_list/<int:price_list_id>/', views.delete_price_list, name='delete_price_list'),
    path('toggle_price_list_status/<int:price_list_id>/', views.toggle_price_list_status, name='toggle_price_list_status'),
    path('add_pricelist_comment/<int:price_list_id>/', views.add_pricelist_comment, name='add_pricelist_comment'),
    path('delete_pricelist_comment/<int:comment_id>/<int:price_list_id>/', views.delete_pricelist_comment, name='delete_pricelist_comment'),
    path('email_pricelist/<int:price_list_id>/', views.email_pricelist, name='email_pricelist'),
    path('whatsapp_pricelist/<int:price_list_id>/', views.whatsapp_pricelist, name='whatsapp_pricelist'),
    path('price_list_pdf/<int:price_list_id>/', views.price_list_pdf, name='price_list_pdf'),
    path('attach_file/<int:price_list_id>/', views.attach_file, name='attach_file'),
    path('import_price_list/', views.import_price_list, name='import_price_list'),
    #----------------------------------------------------------akshay--end--------------------------------------------------------
    
    #-------------------------Arya E.R---------------------------------------------------

    ####  Vendor ###########
    path('vendor',views.vendor,name='vendor'),
    path('view_vendor_list',views.view_vendor_list,name='view_vendor_list'),
    path('add_vendor/',views.add_vendor,name='add_vendor'),
    
    path('view_vendor_active',views.view_vendor_active,name='view_vendor_active'),
    path('view_vendor_inactive',views.view_vendor_inactive,name='view_vendor_inactive'),
    path('sort_vendor_by_name',views.sort_vendor_by_name,name='sort_vendor_by_name'),
    path('sort_vendor_by_amount',views.sort_vendor_by_amount,name='sort_vendor_by_amount'),
    path('delete_vendor/<int:pk>',views.delete_vendor,name='delete_vendor'),
    path('view_vendor_details/<int:pk>',views.view_vendor_details,name='view_vendor_details'),
    path('import_vendor_excel',views.import_vendor_excel,name='import_vendor_excel'),
    path('Vendor_edit/<int:pk>',views.Vendor_edit,name='Vendor_edit'),
    path('do_vendor_edit/<int:pk>',views.do_vendor_edit,name='do_vendor_edit'),
    path('delete_vendors/<int:pk>',views.delete_vendors,name='delete_vendors'),
    path('vendor_status/<int:pk>',views.vendor_status,name='vendor_status'),
    path('vendor_add_comment/<int:pk>',views.vendor_add_comment,name='vendor_add_comment'),
    path('vendor_delete_comment/<int:pk>',views.vendor_delete_comment,name='vendor_delete_comment'),
    path('vendor_shareemail/<int:pk>',views.vendor_shareemail,name='vendor_shareemail'),
    path('payment_terms_add',views.payment_terms_add,name='payment_terms_add'),
    path('add_vendor_file/<int:pk>',views.add_vendor_file,name='add_vendor_file'),

#------------------------------End---------------------------------------------------------

    path('check_term_exist',views.check_term_exist,name='check_term_exist'),
    path('check_email_exist',views.check_email_exist,name='check_email_exist'),
    path('check_work_phone_exist',views.check_work_phone_exist,name='check_work_phone_exist'),
    path('check_phonenumber_exist',views.check_phonenumber_exist,name='check_phonenumber_exist'),
    
    path('check_pan',views.check_pan,name='check_pan'),
    path('check_gst',views.check_gst,name='check_gst'),
    
    path('sort_vendor/<int:selectId>/<int:pk>',views.sort_vendor,name='sort_vendor'),
    path('vendor_status_change/<int:statusId>/<int:pk>',views.vendor_status_change,name='vendor_status_change'),
    
    #---------------- Zoho Final Attendance - Meenu Shaju - Start--------------------
    path('company_attendance_list',views.company_attendance_list,name='company_attendance_list'),
    path('company_mark_attendance',views.company_mark_attendance,name='company_mark_attendance'),
    path('add_attendance',views.add_attendance,name='add_attendance'),
    path('attendance_calendar/<int:employee_id>/<int:target_year>/<int:target_month>/',views.attendance_calendar,name='attendance_calendar'),
    path('attendance_add_comment',views.attendance_add_comment,name='attendance_add_comment'),
    path('attendance_delete_comment/<int:id>',views.delete_attendance_comment,name='attendance_delete_comment'),
    path('attendance_overview/<int:employee_id>/<int:target_month>/<int:target_year>/',views.attendance_overview,name='attendance_overview'),
    path('attendance_pdf/<int:employee_id>/<int:target_month>/<int:target_year>',views.attendance_pdf,name='attendance_pdf'),
    path('attendance_email/<int:employee_id>/<int:target_month>/<int:target_year>',views.attendance_email,name='attendance_email'),
    path('attendance_edit/<int:id>',views.attendance_edit,name='attendance_edit'),
    path('edit_attendance/<int:id>',views.edit_attendance,name='edit_attendance'),
    path('attendance_delete/<int:id>',views.attendance_delete,name='attendance_delete'),
    path('attendance_add_blood',views.attendance_add_blood,name='attendance_add_blood'),
    path('attendance_create_employee',views.attendance_create_employee,name='attendance_create_employee'),
    path('attendance_import',views.attendance_import,name='attendance_import'),
    path('attendance_employee_dropdown',views.attendance_employee_dropdown,name='attendance_employee_dropdown'),
    #---------------- Zoho Final Attendance - Meenu Shaju - End--------------------
    # ------------------------------- GOKUL KRISHNA UR -----------------------------------------

    path('CreateSalaryDetails',views.CreateSalaryDetails,name='CreateSalaryDetails'),
    path('SalaryDetailsListPage',views.SalaryDetailsListPage,name='SalaryDetailsListPage'),
    path('CreateSalaryDetailsFunction',views.CreateSalaryDetailsFunction,name='CreateSalaryDetailsFunction'),
    path('custdata',views.custdata,name='custdata'),
    path('ImportSalaryDetails',views.ImportSalaryDetails,name='ImportSalaryDetails'),

    path('SharePayslipMail/<int:id>',views.SharePayslipMail,name='SharePayslipMail'),
    path('addEmployeeFromSalaryDetails',views.addEmployeeFromSalaryDetails,name='addEmployeeFromSalaryDetails'),
  
    path('addCommentSalaryDetails/<int:id>',views.addCommentSalaryDetails,name='addCommentSalaryDetails'),
    path('DeleteCommentSalaryDetails/<int:id>',views.DeleteCommentSalaryDetails,name='DeleteCommentSalaryDetails'),
    path('EditSalaryDetails/<int:id>',views.EditSalaryDetails,name='EditSalaryDetails'),
    path('SalaryDetailsOverViewPageWithId/<int:id>',views.SalaryDetailsOverViewPageWithId,name='SalaryDetailsOverViewPageWithId'),
    path('SalaryDetailsActiveAndInnactive/<int:id>',views.SalaryDetailsActiveAndInnactive,name='SalaryDetailsActiveAndInnactive'),
    path('EditSalaryDetailsFunction/<int:id>',views.EditSalaryDetailsFunction,name='EditSalaryDetailsFunction'),
    path('SalaryDetailsAddBloodGroup',views.SalaryDetailsAddBloodGroup,name='SalaryDetailsAddBloodGroup'),
    path('SalaryDetailsConvert/<int:id>',views.SalaryDetailsConvert,name='SalaryDetailsConvert'),
    path('SalaryDetailsDelete/<int:id>',views.SalaryDetailsDelete,name='SalaryDetailsDelete'),

# ------------------------------- GOKUL KRISHNA UR -----------------------------------------
    ##--------------------------------emploan by haripriya---------------------------#
    path('Company/employee_listpage',views.employee_listpage,name='employee_listpage'),
    path('Company/employeeloan_create',views.employeeloan_create,name='employeeloan_create'),
    path('Company/check_user_loan',views.check_user_loan,name='check_user_loan'),
    path('Company/listemployee_loan',views.listemployee_loan,name='listemployee_loan'),
    path('Company/addemployeloan',views.addemployeloan,name='addemployeloan'),
    path('Company/employeeloan_details/<int:id>/',views.employeeloan_details, name='employeeloan_details'),
    path('Company/createpayroll2',views.createpayroll2,name='createpayroll2'),
    path('Company/add_emploan_comment/<int:id>/',views.add_emploan_comment, name='add_emploan_comment'),
    path('delete_emploan_comment/<int:ph>/<int:pr>',views.delete_emploan_comment,name='delete_emploan_comment'),
    path('Company/employeeloan_repayment_pageload/<int:id>',views.employeeloan_repayment_pageload,name='employeeloan_repayment_pageload'),
    path('Company/add_repayment/<int:id>/',views.add_repayment, name='add_repayment'),
    path('Company/add_newloan_pageload/<int:id>/',views.add_newloan_pageload, name='add_newloan_pageload'),
    path('Company/add_newloan/<int:id>/',views.add_newloan, name='add_newloan'),
    path('Company/delete_repayment/<int:id>/',views.delete_repayment, name='delete_repayment'),
    
    path('Company/edit_loanrepayment/<int:id>/',views.edit_loanrepayment, name='edit_loanrepayment'),
    path('Company/save_edit_loanrepayment/<int:id>/',views.save_edit_loanrepayment, name='save_edit_loanrepayment'),
    path('Company/edit_additionalloan_pageloage/<int:id>/',views.edit_additionalloan_pageloage, name='edit_additionalloan_pageloage'),
    path('Company/save_edit_additionalloan/<int:id>/',views.save_edit_additionalloan, name='save_edit_additionalloan'),
    path('Company/active_loan/<int:id>/',views.active_loan, name='active_loan'),
    path('Company/inactive_loan/<int:id>/',views.inactive_loan, name='inactive_loan'),
    path('Company/deleteloan/<int:id>/',views.deleteloan, name='deleteloan'),
    path('Company/shareloanemail/<int:pk>/',views.shareloanemail, name='shareloanemail'),
    
    path('Company/edit_loan/<int:id>/',views.edit_loan, name='edit_loan'),  
    path('Company/update_Employeeloan/<int:id>/',views.update_Employeeloan, name='update_Employeeloan'), 
    path('Company/create_loan_duration/',views.create_loan_duration, name='create_loan_duration'), 
    path('Company/loan_duration/',views.loan_duration, name='loan_duration'), 
    path('Company/bankdata/',views.bankdata, name='bankdata'),
    path('Company/bankdata1/',views.bankdata1, name='bankdata1'),
    path('Company/addloan_file/<int:pk>/',views.addloan_file, name='addloan_file'), 
    path('Company/import_employee_loan_details/',views.import_employee_loan_details, name='import_employee_loan_details'),
    #End  
    
    #-----------------------Customer---------------------------#
    #------------Arya E.R---------------#
    path('customer',views.customer,name='customer'),
    path('view_customer_list',views.view_customer_list,name='view_customer_list'),
    path('check_customer_phonenumber_exist',views.check_customer_phonenumber_exist,name='check_customer_phonenumber_exist'),
    path('check_customer_work_phone_exist',views.check_customer_work_phone_exist,name='check_customer_work_phone_exist'),
    path('check_customer_email_exist',views.check_customer_email_exist,name='check_customer_email_exist'),
    path('check_customer_term_exist',views.check_customer_term_exist,name='check_customer_term_exist'),
    path('customer_payment_terms_add',views.customer_payment_terms_add,name='customer_payment_terms_add'),
    path('customer_check_pan',views.customer_check_pan,name='customer_check_pan'),
    path('add_customer/',views.add_customer,name='add_customer'),
    path('customer_check_gst',views.customer_check_gst,name='customer_check_gst'),
    path('sort_customer_by_name',views.sort_customer_by_name,name='sort_customer_by_name'),
    path('sort_customer_by_amount',views.sort_customer_by_amount,name='sort_customer_by_amount'),
    path('view_customer_active',views.view_customer_active,name='view_customer_active'),
    path('view_customer_inactive',views.view_customer_inactive,name='view_customer_inactive'),
    path('import_customer_excel',views.import_customer_excel,name='import_customer_excel'),
    path('view_customer_details/<int:pk>',views.view_customer_details,name='view_customer_details'),
    path('sort_customer/<int:selectId>/<int:pk>',views.sort_customer,name='sort_customer'),
    path('customer_status_change/<int:statusId>/<int:pk>',views.customer_status_change,name='customer_status_change'),
    path('delete_customers/<int:pk>',views.delete_customers,name='delete_customers'),
    path('customer_status/<int:pk>',views.customer_status,name='customer_status'),
    path('customer_add_comment/<int:pk>',views.customer_add_comment,name='customer_add_comment'),
    path('customer_delete_comment/<int:pk>',views.customer_delete_comment,name='customer_delete_comment'), 
    path('add_customer_file/<int:pk>',views.add_customer_file,name='add_customer_file'),
    path('customer_shareemail/<int:pk>',views.customer_shareemail,name='customer_shareemail'),
    path('Customer_edit/<int:pk>',views.Customer_edit,name='Customer_edit'),
    path('do_customer_edit/<int:pk>',views.do_customer_edit,name='do_customer_edit'),
    #---------------------End---------------------------------#
    
    ## kesia loan account ##
    path('zohomodules/loan_account/loan_listing',views.loan_listing,name='loan_listing'),
    path('zohomodules/loan_account/add_loan',views.add_loan,name='add_loan'),
    path('zohomodules/loan_account/holder_dropdown',views.holder_dropdown,name='holder_dropdown'),
    path('zohomodules/loan_account/save_account_details',views.save_account_details,name='save_account_details'),
    path('zohomodules/loan_account/overview/<int:account_id>',views.overview,name='overview'),
    # path('zohomodules/loan_account/transaction/<int:account_id>',views.transaction,name='transaction'),
    path('zohomodules/loan_account/repayment_due_form/<int:account_id>',views.repayment_due_form,name='repayment_due_form'),
    path('zohomodules/loan_account/new_loan/<int:account_id>',views.new_loan,name='new_loan'),
    path('zohomodules/loan_account/update_status/<int:account_id>',views.update_status,name='update_status'),
    path('zohomodules/loan_account/edit_loanaccount/<int:account_id>',views.edit_loanaccount, name='edit_loanaccount'),
    path('zohomodules/loan_account/edit_loantable/<int:account_id>',views.edit_loantable, name='edit_loantable'),
    path('zohomodules/loan_account/edit_repayment/<int:repayment_id>',views.edit_repayment, name='edit_repayment'),
    path('zohomodules/loan_account/edit_additional_loan/<int:repayment_id>',views.edit_additional_loan, name='edit_additional_loan'),
    path('zohomodules/loan_account/share_email/<int:account_id>',views.share_email,name='share_email'),
    path('zohomodules/loan_account/adding_comment/<int:account_id>',views.adding_comment,name='adding_comment'),
    path('zohomodules/loan_account/delete_comment/<int:comment_id>/<int:account_id>',views.delete_comment,name='delete_comment'),
    path('zohomodules/loan_account/get_account_number/<int:account_id>', views.get_account_number, name='get_account_number'),
    path('zohomodules/loan_account/full_account_number/<int:bank_id>', views.full_account_number, name='full_account_number'),
    path('zohomodules/loan_account/delete_repaymenttable/<int:id>',views.delete_repaymenttable,name='delete_repaymenttable'),
    path('zohomodules/loan_account/delete_loan/<int:account_id>',views.delete_loan,name='delete_loan'),
    path('zohomodules/loan_account/statementoverview/<int:account_id>',views.statementoverview,name='statementoverview'),
    #End
    
    path('Company/employeeloan_trans/<int:id>/',views.employeeloan_trans, name='employeeloan_trans'),
    
    path('zohomodules/loan_account/transactoverview/<int:account_id>',views.transactoverview,name='transactoverview'),
    
    path('list_godown/',views.list_godown,name='list_godown'),
    path('add_godown/',views.add_godown,name='add_godown'),
    path('add_godown_func/',views.add_godown_func,name='add_godown_func'),
    path('overview_page/<int:pk>',views.overview_page,name='overview_page'),
    path('edit_godown/<int:pk>',views.edit_godown,name='edit_godown'),
    path('edit_godown_func/',views.edit_godown_func,name='edit_godown_func'),
    path('change_status/<int:pk>',views.change_status,name='change_status'),
    path('change_action/<int:pk>',views.change_action,name='change_action'),
    path('AddComment/<int:pk>',views.AddComment,name='AddComment'),
    path('DeleteComment/<int:pk>',views.DeleteComment,name='DeleteComment'),
    path('AddFile/<int:pk>',views.AddFile,name='AddFile'),
    path('file_download/<int:pk>',views.file_download,name='file_download'),
    path('ShareEmail/<int:pk>',views.ShareEmail,name='ShareEmail'),
    path('Add_Item/',views.Add_Item,name='Add_Item'),
    path('godownmodal_unit/',views.godownmodal_unit,name='godownmodal_unit'),
    path('godownunit_dropdown/',views.godownunit_dropdown,name='godownunit_dropdown'),
    path('AddAccount/',views.AddAccount,name='AddAccount'),
    path('Add_Item_Edit/<int:pk>',views.Add_Item_Edit,name='Add_Item_Edit'),
    path('godownmodal_unit_edit/<int:pk>',views.godownmodal_unit_edit,name='godownmodal_unit_edit'),
    path('Add_Account_Edit/<int:pk>',views.Add_Account_Edit,name='Add_Account_Edit'),
    path('delete_godown/<int:pk>',views.delete_godown,name='delete_godown'),
    
    # -----------------------------------bank_holder--------------------------------sruthi--------------------------
    path('zohomodules/bankholder/list_bankholder',views.list_bankholder, name='list_bankholder'),
    path('bankholder/import_bankholder_excel',views.import_bankholder_excel, name='import_bankholder_excel'),
    path('zohomodules/bankholder/create_holder',views.createholder,name='createholder'),
    path('zohomodules/bankholder/addbnk_refresh',views.addbnk_refresh,name='addbnk_refresh'),
    path('zohomodules/bankholder/create_bankholder',views.create_bankholder,name='create_bankholder'),
    path('bankholder/bankholder_checkname/', views.bankholder_checkname, name='bankholder_checkname'),
    path('bankholder/bankholder_checkacno/', views.bankholder_checkacno, name='bankholder_checkacno'),
    path('zohomodules/bankholder/overview_bankholder/<int:pk>',views.overview_bankholder,name='overview_bankholder'),
    path('zohomodules/bankholder/bankholder_status/<int:pk>',views.bankholder_status,name='bankholder_status'),
    path('zohomodules/bankholder/delete_bankholder/<int:pk>',views.delete_bankholder,name='delete_bankholder'),
    path('zohomodules/bankholder/bankholder_listbank',views.bankholder_listbank,name='bankholder_listbank'),
    path('zohomodules/bankholder/bankholder_checkbank',views.bankholder_checkbank,name='bankholder_checkbank'),
    # path('zohomodules/bankholder/edit_bankholder/bankholder_checkbank',views.bankholder_checkbank,name='bankholder_checkbank'),
    path('zohomodules/bankholder/bankholder_addbank',views.bankholder_addbank,name='bankholder_addbank'),
    path('zohomodules/bankholder/edit_bankholder/<int:pk>',views.edit_bankholder,name='edit_bankholder'),
    path('zohomodules/bankholder/do_bankholder_edit/<int:pk>',views.do_bankholder_edit,name='do_bankholder_edit'),
    path('zohomodules/bankholder/bankholder_checkgst',views.bankholder_checkgst,name='bankholder_checkgst'),
    path('zohomodules/bankholder/bankholder_checkpan',views.bankholder_checkpan,name='bankholder_checkpan'),
    path('zohomodules/bankholder/bankholder_checkemail',views.bankholder_checkemail,name='bankholder_checkemail'),
    path('zohomodules/bankholder/bankholder_checkphone',views.bankholder_checkphone,name='bankholder_checkphone'),
    #End

    path('Company/Payroll/Holiday',views.company_holiday,name='company_holiday'),
    path('Company/Payroll/Holiday/New',views.company_holiday_new,name='company_holiday_new'),
    path('Company/Payroll/Holiday/Import_operation',views.company_holiday_import_operation,name='company_holiday_import_operation'),
    path('Company/Payroll/Holiday/New_add',views.company_holiday_new_add,name='company_holiday_new_add'),
    path('Company/Payroll/Holiday/Overview',views.company_holiday_overview,name='company_holiday_overview'),
    path('Company/Payroll/Holiday/Overview_delete/<int:pk>',views.company_holiday_overview_delete,name='company_holiday_overview_delete'),
    path('Company/Payroll/Holiday/Overview_edit/<int:pk>',views.company_holiday_overview_edit,name='company_holiday_overview_edit'),
    path('Company/Payroll/Holiday/Overview_edit_op/<int:pk>',views.company_holiday_overview_edit_op,name='company_holiday_overview_edit_op'),
    path('Company/Payroll/Holiday/Overview_comment/<int:pk>',views.company_holiday_overview_comment,name='company_holiday_overview_comment'),
    path('Company/Payroll/Holiday/Overview_comment_delete/<int:pk>',views.company_holiday_overview_comment_delete,name='company_holiday_overview_comment_delete'),
    path('Company/Payroll/Holiday/Overview_email_send',views.company_holiday_overview_email_send,name='company_holiday_overview_email_send'),
    path('Company/Payroll/Holiday/import_sample_download',views.company_holiday_import_sample_download,name='company_holiday_import_sample_download'),
    
    #================================manual journal/NEW==============================
    path('manual_journal', views.manual_journal, name='manual_journal'),
    path('add_journal', views.add_journal, name='add_journal'),
    path('journal_overview/<int:journal_id>/', views.journal_overview, name='journal_overview'),
    path('journal',views.journal,name="journal"),
    path('import_journal_list',views.import_journal_list,name='import_journal_list'),
    path('update_journal_status/<int:id>/', views.update_journal_status, name='update_journal_status'),
    path('delete_journal/<int:journal_id>/', views.delete_journal, name='delete_journal'),
    path('add_journal_comment/<int:journal_id>/', views.add_journal_comment, name='add_journal_comment'),
    path('delete_journal_comment/<int:id>/', views.delete_journal_comment, name='delete_journal_comment'),
    path('create_account_jour',views.create_account_jour,name='create_account_jour'),
    path('edit_journal/<int:journal_id>/', views.edit_journal, name='edit_journal'),
    
    path('downloadJournalSampleImportFile',views.downloadJournalSampleImportFile,name='downloadJournalSampleImportFile'),
    path('downloadAccountSampleImportFile',views.downloadAccountSampleImportFile,name='downloadAccountSampleImportFile'),
    path('email_journal/<int:journal_id>/', views.email_journal, name='email_journal'),
    path('journal_pdf/<int:price_list_id>/', views.journal_pdf, name='journal_pdf'),
    path("check_journal_num_valid",views.check_journal_num_valid,name="check_journal_num_valid"),
    
    path("check_journal_num_valid2",views.check_journal_num_valid2,name="check_journal_num_valid2"),
    #End
    
    # -------------------------------Staff section invoice--------------------------------
    path('Staff/invoice_listout',views.invoice_list_out,name='invoice_list_out'),
    path('Staff/invoice/create',views.invoice_create,name='invoice_create'),
    path('Staff/invoice/invoice_createpage',views.invoice_createpage,name='invoice_createpage'),
    path('Staff/invoice/itemdata',views.itemdata,name='itemdata'),
    path('Staff/invoice/viewInvoice',views.viewInvoice,name='viewInvoice'),
    path('Staff/invoice/customerdata',views.customerdata,name='customerdata'),
    path('Staff/invoice/getInvoiceCustomerData',views.getInvoiceCustomerData,name='getInvoiceCustomerData'),
    path('Staff/invoice/getInvItemDetails',views.getInvItemDetails,name='getInvItemDetails'),
    path('Staff/invoice/getBankAccount',views.getBankAccount,name='getBankAccount'),
    path('Staff/invoice/createInvoice',views.createInvoice,name='createInvoice'),
    path('Staff/invoice/checkInvoiceNumber',views.checkInvoiceNumber,name='checkInvoiceNumber'),
    path('Staff/invoice/invoice_import',views.invoice_import,name='invoice_import'),
    path('Staff/invoice/view/<int:pk>',views.view,name='view'),

    path('Staff/invoice/filter_invoice_draft/<int:pk>',views.filter_invoice_draft,name='filter_invoice_draft'),
    path('Staff/invoice/filter_invoice_sent/<int:pk>',views.filter_invoice_sent,name='filter_invoice_sent'),
    path('Staff/invoice/filter_invoice_name/<int:pk>',views.filter_invoice_name,name='filter_invoice_name'),
    path('Staff/invoice/filter_invoice_number/<int:pk>',views.filter_invoice_number,name='filter_invoice_number'),
    path('Staff/invoice/edit_invoice/<int:id>',views.editInvoice, name='editInvoice'),
    path('Staff/invoice/update_invoice/<int:id>',views.updateInvoice, name='updateInvoice'),
    path('Staff/invoice/convert_invoice/<int:id>',views.convertInvoice, name='convertInvoice'),
    path('Staff/invoice/invoicePdf/<int:id>',views.invoicePdf, name='invoicePdf'),
    path('Staff/invoice/delete_invoice/<int:id>',views.deleteInvoice, name= 'deleteInvoice'),
    path('Staff/invoice/Invoice_history/<int:id>',views.InvoiceHistory, name='InvoiceHistory'),
    path('Staff/invoice/add_attach/<int:id>',views.add_attach, name='add_attach'),
    path('Staff/invoice/share_invoice_to_email/<int:id>',views.shareInvoiceToEmail, name='shareInvoiceToEmail'),
    path('Staff/invoice/addInvoiceComment/<int:id>',views.addInvoiceComment, name='addInvoiceComment'),
    path('Staff/invoice/deleteInvoiceComment/<int:id>',views.deleteInvoiceComment, name='deleteInvoiceComment'),
    path('Staff/invoice/newCustomerPaymentTerm',views.newCustomerPaymentTerm, name='newCustomerPaymentTerm'),
    path('Staff/invoice/checkCustomerName',views.checkCustomerName, name='checkCustomerName'),
    path('Staff/invoice/checkCustomerGSTIN',views.checkCustomerGSTIN, name='checkCustomerGSTIN'),
    path('Staff/invoice/checkCustomerPAN',views.checkCustomerPAN, name='checkCustomerPAN'),
    path('Staff/invoice/checkCustomerPhone',views.checkCustomerPhone, name='checkCustomerPhone'),
    path('Staff/invoice/checkCustomerEmail',views.checkCustomerEmail, name='checkCustomerEmail'),
    path('Staff/invoice/createInvoiceCustomer',views.createInvoiceCustomer, name='createInvoiceCustomer'),
    path('Staff/invoice/invoice_item',views.invoice_item, name='invoice_item'),

    path('Staff/invoice/getCustomers',views.getCustomers, name='getCustomers'),
    path('Staff/invoice/createInvoiceItem',views.createInvoiceItem, name='createInvoiceItem'),
    path('Staff/invoice/getItems',views.getItems, name='getItems'),
    path('Staff/invoice/saveItemUnit',views.saveItemUnit, name='saveItemUnit'),
    path('Staff/invoice/show_unit_dropdown',views.show_unit_dropdown, name='show_unit_dropdown'),
    path('Staff/invoice/createNewAccountFromItems',views.createNewAccountFromItems, name='createNewAccountFromItems'),
    path('Staff/invoice/checkAccounts',views.checkAccounts, name='checkAccounts'),
    #End

    # --------------------------------------   ashikhvu   (start)   -----------------------------------------------
    path('recurring_bill_listout',views.recurring_bill_listout,name='recurring_bill_listout'),
    path('recurring_bill_create',views.recurring_bill_create,name='recurring_bill_create'),
    path('get_vendors_details_for_recurr/<int:pk>',views.get_vendors_details_for_recurr,name='get_vendors_details_for_recurr'),
    path('get_customer_details_for_recurr/<int:pk>',views.get_customer_details_for_recurr,name='get_customer_details_for_recurr'),
    path('createReccuringBill',views.createReccuringBill,name='createReccuringBill'),
    path('create_repeat_every',views.create_repeat_every,name='create_repeat_every'),
    path('add_new_creadit_period',views.add_new_creadit_period,name='add_new_creadit_period'),
    # path('createRecurrCustomer',views.createRecurrCustomer,name='createRecurrCustomer'),
    path('check_vendor_work_phone_exist',views.check_vendor_work_phone_exist,name='check_vendor_work_phone_exist'),
    path('check_vendor_phonenumber_exist',views.check_vendor_phonenumber_exist,name='check_vendor_phonenumber_exist'),
    path('recurr_vendor_create',views.recurr_vendor_create,name='recurr_vendor_create'),
    path('recurr_customer_create',views.recurr_customer_create,name='recurr_customer_create'),
    path('add_new_recrring_bill',views.add_new_recrring_bill,name='add_new_recrring_bill'),
    path('recurr_overview/<int:pk>',views.recurr_overview,name='recurr_overview'),
    path('check_rec_bill_no_valid',views.check_rec_bill_no_valid,name='check_rec_bill_no_valid'),
    path('get_price_list_percentage/<int:pk>',views.get_price_list_percentage,name='get_price_list_percentage'),
    # --------------------------------------   ashikhvu   (end)   -----------------------------------------------


    re_path(r'^static/(?P<path>.*)$', serve,{'document_root': settings.STATIC_ROOT}),
]
urlpatterns += static(settings.STATIC_URL, document_root=settings.STATICFILES_DIRS)
urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)