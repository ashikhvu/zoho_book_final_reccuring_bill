#Zoho Final
from django.db import models
from Register_Login.models import *
from django.contrib.auth.models import User
from Register_Login.models import LoginDetails,CompanyDetails
from Register_Login.models import LoginDetails,CompanyDetails,Company_Payment_Term
from datetime import datetime
from datetime import date
# Create your models here.

#---------------- models for zoho modules--------------------
# TINTO -----ITEM ----START

class Unit(models.Model):
 
    unit_name=models.CharField(max_length=255)
    company=models.ForeignKey(CompanyDetails,on_delete=models.CASCADE)


class Items(models.Model):
   
    item_type=models.CharField(max_length=255)
    item_name=models.CharField(max_length=255)
   
    unit=models.ForeignKey(Unit,on_delete=models.CASCADE)
    hsn_code=models.IntegerField(null=True,blank=True)
    tax_reference=models.CharField(max_length=255,null=True)
    intrastate_tax=models.IntegerField(null=True,blank=True)
    interstate_tax=models.IntegerField(null=True,blank=True)

    selling_price=models.IntegerField(null=True,blank=True)
    sales_account=models.CharField(max_length=255)
    sales_description=models.CharField(max_length=255)

    purchase_price=models.IntegerField(null=True,blank=True)
    purchase_account=models.CharField(max_length=255)
    purchase_description=models.CharField(max_length=255)
   
    minimum_stock_to_maintain=models.IntegerField(blank=True,null=True)  
    activation_tag=models.CharField(max_length=255,default='active')
    inventory_account=models.CharField(max_length=255,null=True)

    date=models.DateTimeField(auto_now_add=True)                                       

    opening_stock=models.IntegerField(blank=True,null=True,default=0)
    current_stock=models.IntegerField(blank=True,null=True,default=0)
    opening_stock_per_unit=models.IntegerField(blank=True,null=True,)
    company=models.ForeignKey(CompanyDetails,on_delete=models.CASCADE)
    login_details=models.ForeignKey(LoginDetails,on_delete=models.CASCADE)

    type=models.CharField(max_length=255,blank=True,null=True)

    track_inventory=models.IntegerField(blank=True,null=True)

class Item_Transaction_History(models.Model):
    company=models.ForeignKey(CompanyDetails,on_delete=models.CASCADE)
    logindetails=models.ForeignKey(LoginDetails,on_delete=models.CASCADE)
    items=models.ForeignKey(Items,on_delete=models.CASCADE)
    Date=models.DateField(null=True)
    action=models.CharField(max_length=255,default='Created')

class Items_comments(models.Model):                                              
    company=models.ForeignKey(CompanyDetails,on_delete=models.CASCADE)
    logindetails=models.ForeignKey(LoginDetails,on_delete=models.CASCADE)
    Items=models.ForeignKey(Items,on_delete=models.CASCADE)
    comments = models.CharField(max_length=255,null=True,blank=True)


# TINTO -----ITEM ----END
    
# TINTO -----CHART OF ACCOUNNTS ----START
    
class Chart_of_Accounts(models.Model):
  
    account_type = models.CharField(max_length=255,null=True,blank=True)
    account_name = models.CharField(max_length=255,null=True,blank=True)

    account_description = models.CharField(max_length=255,null=True,blank=True)

    account_number = models.CharField(max_length=255,null=True,blank=True)
    
    account_code = models.CharField(max_length=255,null=True,blank=True)
    description = models.CharField(max_length=255,null=True,blank=True)
    status=models.CharField(max_length=255,null=True,blank=True,default='Active')
    Create_status = models.CharField(max_length=255,null=True,blank=True,default='added')
    company=models.ForeignKey(CompanyDetails,on_delete=models.CASCADE)
    login_details=models.ForeignKey(LoginDetails,on_delete=models.CASCADE)
    sub_account = models.CharField(max_length=255,null=True,blank=True)
    parent_account = models.CharField(max_length=255,null=True,blank=True)

class Chart_of_Accounts_History(models.Model):
    company=models.ForeignKey(CompanyDetails,on_delete=models.CASCADE)
    logindetails=models.ForeignKey(LoginDetails,on_delete=models.CASCADE)
    chart_of_accounts=models.ForeignKey(Chart_of_Accounts,on_delete=models.CASCADE)
    Date=models.DateField(null=True)
    action=models.CharField(max_length=255,default='Created')



class chart_of_accounts_comments(models.Model):                                         
    company=models.ForeignKey(CompanyDetails,on_delete=models.CASCADE)
    logindetails=models.ForeignKey(LoginDetails,on_delete=models.CASCADE)
    chart_of_accounts=models.ForeignKey(Chart_of_Accounts,on_delete=models.CASCADE)
    comments = models.CharField(max_length=255,null=True,blank=True)
    
# TINTO -----CHART OF ACCOUNNTS ----END


#--------------------------GEORGE MATHEW____________
class payroll_employee(models.Model):
    title = models.CharField(max_length=100,null=True)
    first_name = models.CharField(max_length=100,null=True)
    last_name = models.CharField(max_length=100,null=True)
    alias = models.CharField(max_length=100,null=True)
    image=models.ImageField(upload_to="image/", null=True)
    joindate=models.DateField(null=True)
    salary_type = models.CharField(max_length=100, default='Fixed',null=True)
    salary = models.IntegerField(null=True,blank=True)
    emp_number = models.CharField(max_length=100,null=True)
    designation = models.CharField(max_length=100,null=True)
    location = models.CharField(max_length=100,null=True)
    gender = models.CharField(max_length=100,null=True)
    dob=models.DateField(null=True)
    age = models.PositiveIntegerField(default=0)
    blood = models.CharField(max_length=10,null=True)
    parent = models.CharField(max_length=100,null=True)
    spouse_name = models.CharField(max_length=100,null=True)
    address = models.CharField(max_length=250,null=True)
    permanent_address = models.CharField(max_length=250,null=True)
    Phone = models.BigIntegerField(null=True)
    emergency_phone = models.BigIntegerField(null=True ,blank=True,default=1)
    email = models.EmailField(max_length=255,null=True)
    Income_tax_no = models.CharField(max_length=255,null=True)
    Aadhar = models.CharField(max_length=250,default='',null=True)
    UAN = models.CharField(max_length=255,null=True)
    PFN = models.CharField(max_length=255,null=True)
    PRAN = models.CharField(max_length=255,null=True)
    status=models.CharField(max_length=200,default='Active',null=True)
    isTDS=models.CharField(max_length=200,null=True)
    TDS_percentage = models.IntegerField(null=True,default=0)
    salaryrange = models.CharField(max_length=10, choices=[('1-10', '1-10'), ('10-15', '10-15'), ('15-31', '15-31')], default='1-10',null=True)
    amountperhr = models.IntegerField(default=0,blank=True,null=True)
    workhr = models.IntegerField(default=0,blank=True,null=True)
    uploaded_file=models.FileField(upload_to="images/",null=True)
    company=models.ForeignKey(CompanyDetails,on_delete=models.CASCADE,null=True)
    login_details=models.ForeignKey(LoginDetails,on_delete=models.CASCADE,null=True)
    acc_no = models.CharField(null=True,max_length=255)  
    IFSC = models.CharField(max_length=100,null=True)
    bank_name = models.CharField(max_length=100,null=True)
    branch = models.CharField(max_length=100,null=True)
    transaction_type = models.CharField(max_length=100,null=True)
    
class employee_history(models.Model):
    company=models.ForeignKey(CompanyDetails,on_delete=models.CASCADE,null=True)
    login_details=models.ForeignKey(LoginDetails,on_delete=models.CASCADE,null=True)
    employee=models.ForeignKey(payroll_employee,on_delete=models.CASCADE,null=True)
    Date=models.DateField(null=True,auto_now=True)
    Action=models.CharField(null=True,max_length=255)
    
class Bloodgroup(models.Model):
    Blood_group=models.CharField(max_length=255,null=True)
    
class comment(models.Model):
    comment=models.CharField(null=True,max_length=255)
    login_details=models.ForeignKey(LoginDetails,on_delete=models.CASCADE,null=True)
    employee=models.ForeignKey(payroll_employee,on_delete=models.CASCADE,null=True)
#------------------------------------------------------------------end-------------------------------------------------------


class payroll_employee_comment(models.Model):
    comment=models.CharField(null=True,max_length=255)
    login_details=models.ForeignKey(LoginDetails,on_delete=models.CASCADE,null=True)
    employee=models.ForeignKey(payroll_employee,on_delete=models.CASCADE,null=True)
    
    
#----------------- Banking -----------------------------#

class Banking(models.Model):
    login_details = models.ForeignKey(LoginDetails, on_delete=models.CASCADE,null=True,blank=True)
    company=models.ForeignKey(CompanyDetails,on_delete=models.CASCADE)
    bnk_name = models.CharField(max_length=220,default='', null=True, blank=True)
    bnk_branch = models.CharField(max_length=220,default='', null=True, blank=True)
    bnk_acno = models.CharField(max_length=220,default='', null=True, blank=True)
    bnk_ifsc = models.CharField(max_length=220,default='', null=True, blank=True)
    BAL_TYPE = [
        ('Credit', 'Credit'),
        ('Debit', 'Debit'),
    ]
    bnk_bal_type = models.CharField(max_length=220,choices=BAL_TYPE, default='Debit')
    bnk_opnbal =models.FloatField(null=True, blank=True)
    bnk_bal =models.FloatField(null=True, blank=True)
    date = models.DateField(null=True, blank=True)
    document=models.FileField(upload_to='bank/',null=True,blank=True)
    status= models.TextField(default='Active')

 
class BankTransaction(models.Model):
    login_details = models.ForeignKey(LoginDetails, on_delete=models.CASCADE,null=True,blank=True)
    company = models.ForeignKey(CompanyDetails,on_delete=models.CASCADE)
    banking = models.ForeignKey(Banking,on_delete=models.CASCADE)
    trans_cur_amount = models.FloatField(null=True, blank=True)
    trans_amount = models.FloatField(null=True, blank=True)
    trans_adj_amount = models.FloatField(null=True, blank=True)
    trans_adj_date = models.DateField(null=True, blank=True)

    TRANS_TYPE = [
        ('Opening Balance', 'Opening Balance'),
        ('Bank to Bank', 'Bank to Bank'),
        ('Bank to Cash', 'Bank to Cash'),
        ('Cash to Bank', 'Cash to Bank'),
        ('Bank Adjustment', 'Bank Adjustment'),
    ]
    trans_type = models.CharField(max_length=220,choices=TRANS_TYPE)

    ADJ_TYPE = [
        ('', ''),
        ('Balance Increase', 'Balance Increase'),
        ('Balance Decrease', 'Balance Decrease'),
    ]
    trans_adj_type = models.CharField(max_length=220,choices=ADJ_TYPE)
    trans_desc = models.CharField(max_length=220,null=True,blank=True)
    bank_to_bank_no = models.PositiveIntegerField(null=True,blank=True)


class BankingHistory(models.Model):
    login_details = models.ForeignKey(LoginDetails, on_delete=models.CASCADE,null=True,blank=True)
    company = models.ForeignKey(CompanyDetails,on_delete=models.CASCADE)
    banking = models.ForeignKey(Banking,on_delete=models.CASCADE)
    hist_adj_amount = models.FloatField(null=True, blank=True)
    hist_adj_date = models.DateField(auto_now_add=True, null=True, blank=True)
    ACTION_TYPE = [
        ('Created', 'Created'),
        ('Updated', 'Updated'),
    ]
    hist_action = models.CharField(max_length=220,choices=ACTION_TYPE)

class BankTransactionHistory(models.Model):
    login_details = models.ForeignKey(LoginDetails, on_delete=models.CASCADE,null=True,blank=True)
    company = models.ForeignKey(CompanyDetails,on_delete=models.CASCADE)
    transaction = models.ForeignKey(BankTransaction,on_delete=models.CASCADE,null=True,blank=True)
    hist_cur_amount = models.FloatField(null=True, blank=True)
    hist_amount = models.FloatField(null=True, blank=True)
    hist_adj_amount = models.FloatField(null=True, blank=True)
    hist_adj_date = models.DateField(auto_now_add=True, null=True, blank=True)
    ACTION_TYPE = [
        ('Created', 'Created'),
        ('Updated', 'Updated'),
    ]
    hist_action = models.CharField(max_length=220,choices=ACTION_TYPE)
    
    
#----------------------------------------------------------akshay--start--------------------------------------------------------


#------------------- PRICE LIST MODULE ------------

class PriceList(models.Model):
    
    name = models.CharField(max_length=255, null=True)
    type_choices = [
        ('Sales', 'Sales'),('Purchase', 'Purchase'),]
    type = models.CharField(max_length=10, choices=type_choices, null=True)
    item_rate_choices = [('Percentage', 'Percentage'),('Each Item', 'Each Item'),]
    item_rate_type = models.CharField(max_length=15, choices=item_rate_choices, null=True)
    description = models.TextField(null=True)
    percentage_type_choices = [('Markup', 'Markup'),('Markdown', 'Markdown'),]
    percentage_type = models.CharField(max_length=10, choices=percentage_type_choices, null=True, blank=True)
    percentage_value = models.IntegerField(null=True, blank=True)
    round_off_choices = [
        ('Never Mind', 'Never Mind'),
        ('Nearest Whole Number', 'Nearest Whole Number'),
        ('0.99', '0.99'),
        ('0.50', '0.50'),
        ('0.49', '0.49'),
    ]
    round_off = models.CharField(max_length=20, choices=round_off_choices, null=True)
    currency_choices = [('Indian Rupee', 'Indian Rupee')]
    currency = models.CharField(max_length=20, choices=currency_choices, null=True)
    date = models.DateField(auto_now_add=True, null=True)
    STATUS_CHOICES = [('Active', 'Active'),('Inactive', 'Inactive'),]
    status = models.CharField(max_length=10, choices=STATUS_CHOICES, default='Active')
    attachment = models.FileField(upload_to='price_list_attachment/', null=True, blank=True)

    company = models.ForeignKey(CompanyDetails, on_delete=models.CASCADE)
    login_details = models.ForeignKey(LoginDetails, on_delete=models.CASCADE)

class PriceListItem(models.Model):
    company = models.ForeignKey(CompanyDetails, on_delete=models.CASCADE)
    login_details = models.ForeignKey(LoginDetails, on_delete=models.CASCADE)
    price_list = models.ForeignKey(PriceList, on_delete=models.CASCADE)
    item = models.ForeignKey(Items, on_delete=models.CASCADE)  
    standard_rate = models.DecimalField(max_digits=10, decimal_places=2,null=True,blank=True)
    custom_rate = models.DecimalField(max_digits=10, decimal_places=2,null=True,blank=True)

class PriceListTransactionHistory(models.Model):
    company = models.ForeignKey(CompanyDetails, on_delete=models.CASCADE)
    login_details = models.ForeignKey(LoginDetails, on_delete=models.CASCADE)
    price_list = models.ForeignKey(PriceList, on_delete=models.CASCADE)
    date = models.DateField(auto_now_add=True,null=True)
    action_choices = [
        ('Created', 'Created'), 
        ('Edited', 'Edited')
        ]
    action = models.CharField(max_length=10, choices=action_choices,null=True)

class PriceListComment(models.Model):
    company = models.ForeignKey(CompanyDetails, on_delete=models.CASCADE)
    login_details = models.ForeignKey(LoginDetails, on_delete=models.CASCADE)
    price_list = models.ForeignKey(PriceList, on_delete=models.CASCADE)
    comment = models.TextField()
    date_added = models.DateField(auto_now_add=True)

#----------------------------------------------------------akshay--end--------------------------------------------------------

#-----------------Arya E.R----------------------------------------

class Vendor(models.Model):
    title = models.CharField(max_length=255,null=True,blank=True)
    first_name = models.CharField(max_length=255,null=True,blank=True)
    last_name = models.CharField(max_length=255,null=True,blank=True)
    vendor_display_name = models.CharField(max_length=255,null=True,blank=True)
    vendor_email = models.EmailField()
    mobile = models.CharField(max_length=15,default='')
    phone = models.CharField(max_length=15,default='')
    company_name = models.CharField(max_length=255,null=True,blank=True)
    skype_name_number = models.CharField(max_length=255,null=True,blank=True)
    designation = models.CharField(max_length=255,null=True,blank=True)
    department = models.CharField(max_length=255,null=True,blank=True)
    website = models.URLField(blank=True, null=True,default='')
    gst_treatment = models.CharField(max_length=255,null=True,blank=True)
    gst_number = models.CharField(max_length=20,null=True,blank=True)
    pan_number = models.CharField(max_length=20,null=True,blank=True)
    currency = models.CharField(max_length=255,null=True,blank=True)
    opening_balance_type = models.CharField(max_length=255,null=True,blank=True)
    opening_balance = models.DecimalField(max_digits=10, decimal_places=2,default=0.00)
    current_balance = models.DecimalField(max_digits=10, decimal_places=2,default=0.00)
    credit_limit = models.DecimalField(max_digits=10, decimal_places=2,default=0.00)
    source_of_supply = models.CharField(max_length=255,null=True,blank=True)
    payment_term = models.ForeignKey(Company_Payment_Term, on_delete=models.SET_NULL,null=True,blank=True)
    billing_attention = models.CharField(max_length=255,null=True,blank=True)
    billing_address = models.TextField(null=True,blank=True)
    billing_city = models.CharField(max_length=255,null=True,blank=True)
    billing_state = models.CharField(max_length=255,null=True,blank=True)
    billing_country = models.CharField(max_length=255,null=True,blank=True)
    billing_pin_code = models.CharField(max_length=10,null=True,blank=True)
    billing_phone = models.CharField(max_length=15,null=True,blank=True)
    billing_fax = models.CharField(max_length=15,null=True,blank=True)
    shipping_attention = models.CharField(max_length=255,null=True,blank=True)
    shipping_address = models.TextField(null=True,blank=True)
    shipping_city = models.CharField(max_length=255,null=True,blank=True)
    shipping_state = models.CharField(max_length=255,null=True,blank=True)
    shipping_country = models.CharField(max_length=255,null=True,blank=True)
    shipping_pin_code = models.CharField(max_length=10,null=True,blank=True)
    shipping_phone = models.CharField(max_length=15,null=True,blank=True)
    shipping_fax = models.CharField(max_length=15,null=True,blank=True)
    remarks = models.TextField(null=True,blank=True)
    vendor_status = models.CharField(max_length=10,null=True,blank=True)
    company = models.ForeignKey(CompanyDetails, on_delete=models.CASCADE,null=True,blank=True)
    login_details = models.ForeignKey(LoginDetails, on_delete=models.CASCADE,null=True,blank=True)
    
    def __str__(self):
        return f"{self.first_name} {self.last_name}"

class VendorContactPerson(models.Model):
    company = models.ForeignKey(CompanyDetails, on_delete=models.CASCADE)
    vendor = models.ForeignKey(Vendor, on_delete=models.CASCADE)
    title = models.CharField(max_length=255)
    first_name = models.CharField(max_length=255)
    last_name = models.CharField(max_length=255)
    email = models.EmailField()
    work_phone = models.CharField(max_length=15)
    mobile = models.CharField(max_length=15)
    skype_name_number = models.CharField(max_length=255)
    designation = models.CharField(max_length=255)
    department = models.CharField(max_length=255)

    def __str__(self):
        return f"{self.first_name} {self.last_name}"

class VendorHistory(models.Model):
    company = models.ForeignKey(CompanyDetails, on_delete=models.CASCADE)
    login_details = models.ForeignKey(LoginDetails, on_delete=models.CASCADE)
    vendor = models.ForeignKey(Vendor, on_delete=models.CASCADE)
    date = models.DateField()
    action = models.CharField(max_length=200,null=True,blank=True)

    def __str__(self):
        return f"{self.vendor} - {self.action}"
    
class Vendor_remarks_table(models.Model):
    company = models.ForeignKey(CompanyDetails, on_delete=models.CASCADE)
    vendor=models.ForeignKey(Vendor,on_delete=models.CASCADE,null=True)
    remarks=models.CharField(max_length=500)

class Vendor_comments_table(models.Model):
    login_details=models.ForeignKey(LoginDetails,on_delete=models.CASCADE,null=True)
    company = models.ForeignKey(CompanyDetails, on_delete=models.CASCADE)
    vendor=models.ForeignKey(Vendor,on_delete=models.CASCADE,null=True)
    comment=models.TextField(max_length=500)

class Vendor_mail_table(models.Model):
    company = models.ForeignKey(CompanyDetails, on_delete=models.CASCADE)
    vendor=models.ForeignKey(Vendor,on_delete=models.CASCADE,null=True)
    mail_from=models.TextField(max_length=300)
    mail_to=models.TextField(max_length=300)
    subject=models.TextField(max_length=250)
    content=models.TextField(max_length=900)
    mail_date=models.DateTimeField(auto_now_add=True)

class Vendor_doc_upload_table(models.Model):
    login_details=models.ForeignKey(LoginDetails,on_delete=models.CASCADE,null=True)
    company = models.ForeignKey(CompanyDetails, on_delete=models.CASCADE)
    vendor=models.ForeignKey(Vendor,on_delete=models.CASCADE,null=True)
    title=models.TextField(max_length=200)
    document=models.FileField(upload_to='doc/')

#--------------------------------------end-----------------------------------------------------------

class Holiday(models.Model):
    start_date = models.DateField(null=True, blank=True)
    end_date = models.DateField(null=True, blank=True)
    holiday_name = models.CharField(max_length=255, null=True, blank=True)
    user = models.ForeignKey(LoginDetails, on_delete=models.CASCADE, null=True, blank=True)
    company=models.ForeignKey(CompanyDetails, on_delete=models.CASCADE, null=True,blank=True)
    
class CompanyRepeatEvery(models.Model):
    company = models.ForeignKey(CompanyDetails, on_delete=models.CASCADE,null=True,blank=True)
    repeat_every =models.CharField(max_length=100,null=True,blank=True,default='')
    repeat_type =models.CharField(max_length=100,null=True,blank=True,default='')
    duration =models.IntegerField(null=True,default=0)
    days =models.IntegerField(null=True,default=0)
    
    
#---------------- Zoho Final Attendance - Meenu Shaju - Start--------------------

class Attendance(models.Model):
    company=models.ForeignKey(CompanyDetails,on_delete=models.CASCADE,null=True)
    login_details=models.ForeignKey(LoginDetails,on_delete=models.CASCADE,null=True)
    employee=models.ForeignKey(payroll_employee,on_delete=models.CASCADE,null=True)
    holiday=models.ForeignKey(Holiday,on_delete=models.CASCADE,null=True)
    date=models.DateField(null=True)
    status=models.CharField(max_length=255,null=True)
    reason=models.CharField(max_length=255,null=True)

    
class Attendance_History(models.Model):
    company=models.ForeignKey(CompanyDetails,on_delete=models.CASCADE,null=True)
    login_details=models.ForeignKey(LoginDetails,on_delete=models.CASCADE,null=True)
    attendance=models.ForeignKey(Attendance,on_delete=models.CASCADE,null=True)
    date=models.DateField(null=True)
    action=models.CharField(max_length=100,null=True)

class Attendance_comment(models.Model):
    company=models.ForeignKey(CompanyDetails,on_delete=models.CASCADE,null=True)
    login_details=models.ForeignKey(LoginDetails,on_delete=models.CASCADE,null=True)
    employee=models.ForeignKey(payroll_employee,on_delete=models.CASCADE,null=True)
    comment = models.TextField(null=True) 
    month = models.IntegerField(null=True)  
    year = models.IntegerField(null=True)  

#---------------- Zoho Final Attendance - Meenu Shaju - End--------------------


# ------------------------------- GOKUL KRISHNA UR -----------------------------------------

class SalaryDetails(models.Model):
    company=models.ForeignKey(CompanyDetails,on_delete=models.CASCADE,null=True)
    employee=models.ForeignKey(payroll_employee,on_delete=models.CASCADE,null=True)
    attendance=models.ForeignKey(Attendance,on_delete=models.CASCADE,null=True)
    holiday=models.IntegerField(default=0,blank=True,null=True)
    salary_date =models.DateField(null=True)
    casual_leave = models.IntegerField(default=0,blank=True,null=True)
    month =  models.CharField(max_length=100,null=True)
    year = models.IntegerField(default=0,blank=True,null=True)
    basic_salary = models.IntegerField(default=0,blank=True,null=True)
    conveyance_allowance = models.IntegerField(default=0,blank=True,null=True)
    hra = models.IntegerField(default=0,blank=True,null=True)
    other_allowance = models.IntegerField(default=0,blank=True,null=True)
    total_working_days = models.IntegerField(default=0,blank=True,null=True)
    other_cuttings = models.IntegerField(default=0,blank=True,null=True)
    add_bonus = models.IntegerField(default=0,blank=True,null=True)
    salary = models.FloatField(default=0,blank=True,null=True)
    description = models.CharField(max_length=100,null=True)
    status = models.CharField(max_length=100,null=True,default='Active')
    DraftorSave = models.CharField(max_length=100,null=True)
    total_amount= models.FloatField(default=0,blank=True,null=True)
    

class CommentSalaryDetails(models.Model):
    employee=models.ForeignKey(payroll_employee,on_delete=models.CASCADE,null=True)
    comment=models.CharField(max_length=100,null=True)
    salary_details = models.ForeignKey(SalaryDetails,on_delete=models.CASCADE,null=True)
    company=models.ForeignKey(CompanyDetails,on_delete=models.CASCADE,null=True)
    

class HistorySalaryDetails(models.Model):
    company=models.ForeignKey(CompanyDetails,on_delete=models.CASCADE,null=True)
    login_details=models.ForeignKey(LoginDetails,on_delete=models.CASCADE,null=True)
    salary_details = models.ForeignKey(SalaryDetails,on_delete=models.CASCADE,null=True)
    date = models.DateField(auto_now_add=True)
    ADD = 'add'
    EDIT = 'edit'
    ACTION_CHOICES = [
        (ADD, 'Add'),
        (EDIT, 'Edit'), 
    ]
    action = models.CharField(max_length=7, choices=ACTION_CHOICES, default=ADD)

# ------------------------------- GOKUL KRISHNA UR -----------------------------------------

#---------------------EMPLOYEE_LOAN------------------------------------------#by haripriya

class LoanDuration(models.Model):
    company=models.ForeignKey(CompanyDetails,on_delete=models.CASCADE)
    logindetails=models.ForeignKey(LoginDetails,on_delete=models.CASCADE)
    day = models.IntegerField(null=True, blank=True)
    duration = models.CharField(max_length=50, choices=(
        ('Months', 'Months'),
        ('Month', 'Month'),
        ('Years', 'Years'),
        ('Year', 'Year'),
    ))

class EmployeeLoan(models.Model):
    Employee = models.ForeignKey(payroll_employee,on_delete=models.CASCADE,null=True,blank=True)
    
    Loandate = models.DateField(null=True)
    LoanAmount =  models.IntegerField(null=True, blank=True)
    duration = models.CharField(max_length=255, blank=True)
    Expiry_date = models.DateField(null=True)
    payment_method = models.CharField(max_length=220,null=True,blank=True)
    cheque_number = models.CharField(max_length=220,null=True,blank=True)
    upi_id =models.CharField(max_length=220,null=True,blank=True)
    bank_acc_number =models.CharField(max_length=220,null=True,blank=True)
    Monthly_payment_type =models.CharField(max_length=220,null=True,blank=True)
    MonthlyCut_percentage = models.IntegerField(null=True,blank=True)
    MonthlyCut_Amount =models.IntegerField(null=True,blank=True)
    note = models.CharField(max_length=220,null=True,blank=True)
    file = models.FileField(upload_to="images/",null=True)
    status =models.CharField(max_length=200,null=True,blank=True,default='')
    login_details = models.ForeignKey(LoginDetails, on_delete=models.CASCADE,null=True,blank=True)
    company = models.ForeignKey(CompanyDetails,on_delete=models.CASCADE,null=True,blank=True)
    balance=models.IntegerField(null=True,blank=True)
    active = models.BooleanField(default=True)
    emp_name= models.CharField(max_length=220,null=True,blank=True)
    emp_no= models.IntegerField(null=True,blank=True)
    join_date = models.DateField(null=True)
    salary = models.IntegerField(null=True,blank=True)
    email= models.EmailField(max_length=255,null=True)
 

class Employeeloan_history(models.Model):
    login_details = models.ForeignKey(LoginDetails, on_delete=models.CASCADE,null=True,blank=True)
    company = models.ForeignKey(CompanyDetails,on_delete=models.CASCADE,null=True,blank=True)
    employeeloan =models.ForeignKey(EmployeeLoan,on_delete=models.CASCADE,null=True,blank=True)
    
    Date = models.DateField(null=True,auto_now=True)
    action = models.CharField(max_length=220,null=True,blank=True)

class employeeloan_comments(models.Model):                                         
    company=models.ForeignKey(CompanyDetails,on_delete=models.CASCADE)
    logindetails=models.ForeignKey(LoginDetails,on_delete=models.CASCADE)
    employee=models.ForeignKey(EmployeeLoan,on_delete=models.CASCADE)
    comments = models.CharField(max_length=255,null=True,blank=True)

class EmployeeLoanRepayment(models.Model):
    employee = models.ForeignKey(payroll_employee, on_delete=models.CASCADE, null=True)
    principal_amount = models.IntegerField(null=True)
    interest_amonut = models.IntegerField(null=True)
    payment_date = models.DateField(null=True)
    payment_method = models.CharField(max_length=255,null=True)
    cheque_id=models.CharField(null=True,blank=True,max_length=255)
    upi_id=models.CharField(null=True,blank=True,max_length=255)
    bank_id=models.CharField(null=True,blank=True,max_length=255)
    total_payment = models.IntegerField(null=True)
    balance = models.IntegerField(null=True)
    particular = models.CharField(max_length=255,null=True)
    emp = models.ForeignKey(EmployeeLoan, on_delete=models.CASCADE, null=True)
    company=models.ForeignKey(CompanyDetails,on_delete=models.CASCADE)
    logindetails=models.ForeignKey(LoginDetails,on_delete=models.CASCADE)

#..........................Employeeloan end...........................#

class Customer(models.Model):
    login_details = models.ForeignKey(LoginDetails, on_delete=models.CASCADE,null=True,blank=True)
    company = models.ForeignKey(CompanyDetails,on_delete=models.CASCADE,null=True,blank=True)
    company_payment_terms = models.ForeignKey(Company_Payment_Term,on_delete=models.CASCADE,null=True,blank=True)

    customer_type = models.CharField(max_length=220,null=True,blank=True)
    title = models.CharField(max_length=220,null=True,blank=True)
    first_name = models.CharField(max_length=220,null=True,blank=True)
    last_name = models.CharField(max_length=220,null=True,blank=True)
    customer_display_name = models.CharField(max_length=220,null=True,blank=True)
    company_name = models.CharField(max_length=220,null=True,blank=True)
    customer_email = models.EmailField(max_length=255,null=True)
    customer_phone = models.CharField(max_length=220,null=True,blank=True)
    customer_mobile = models.CharField(max_length=220,null=True,blank=True)

    skype = models.CharField(max_length=220,null=True,blank=True)
    designation = models.CharField(max_length=220,null=True,blank=True)
    department = models.CharField(max_length=220,null=True,blank=True)
    website = models.CharField(max_length=220,null=True,blank=True)
    GST_treatement = models.CharField(max_length=220,null=True,blank=True)
    GST_number = models.CharField(max_length=220,null=True,blank=True)
    PAN_number = models.CharField(max_length=220,null=True,blank=True)
    place_of_supply = models.CharField(max_length=220,null=True,blank=True)
    tax_preference = models.CharField(max_length=220,null=True,blank=True)

    currency = models.CharField(max_length=220,null=True,blank=True)
    opening_balance_type = models.CharField(max_length=220,null=True,blank=True)
    opening_balance = models.FloatField(null=True, blank=True,default=0.00)
    credit_limit = models.FloatField(null=True, blank=True)
    price_list = models.CharField(max_length=220,null=True,blank=True)
    portal_language = models.CharField(max_length=220,null=True,blank=True)

    facebook = models.CharField(max_length=220,null=True,blank=True)
    twitter = models.CharField(max_length=220,null=True,blank=True)
    current_balance = models.FloatField(null=True, blank=True,default=0.00)

    billing_attention = models.CharField(max_length=220,null=True,blank=True)
    billing_address = models.CharField(max_length=220,null=True,blank=True)
    billing_city = models.CharField(max_length=220,null=True,blank=True)
    billing_state = models.CharField(max_length=220,null=True,blank=True)
    billing_country = models.CharField(max_length=220,null=True,blank=True)
    billing_pincode = models.CharField(max_length=220,null=True,blank=True)
    billing_mobile = models.CharField(max_length=220,null=True,blank=True)
    billing_fax = models.CharField(max_length=220,null=True,blank=True)

    shipping_attention = models.CharField(max_length=220,null=True,blank=True)
    shipping_address = models.CharField(max_length=220,null=True,blank=True)
    shipping_city = models.CharField(max_length=220,null=True,blank=True)
    shipping_state = models.CharField(max_length=220,null=True,blank=True)
    shipping_country = models.CharField(max_length=220,null=True,blank=True)
    shipping_pincode = models.CharField(max_length=220,null=True,blank=True)
    shipping_mobile = models.CharField(max_length=220,null=True,blank=True)
    shipping_fax = models.CharField(max_length=220,null=True,blank=True)

    remarks = models.CharField(max_length=220,null=True,blank=True)
    customer_status = models.CharField(max_length=220,null=True,blank=True)


class Customer_remarks_table(models.Model):
    company = models.ForeignKey(CompanyDetails, on_delete=models.CASCADE)
    customer=models.ForeignKey(Customer,on_delete=models.CASCADE,null=True)
    remarks=models.CharField(max_length=500)   

class Customer_comments_table(models.Model):
    login_details=models.ForeignKey(LoginDetails,on_delete=models.CASCADE,null=True)
    company = models.ForeignKey(CompanyDetails, on_delete=models.CASCADE)
    customer=models.ForeignKey(Customer,on_delete=models.CASCADE,null=True)
    comment=models.TextField(max_length=500)  

class Customer_doc_upload_table(models.Model):
    login_details=models.ForeignKey(LoginDetails,on_delete=models.CASCADE,null=True)
    company = models.ForeignKey(CompanyDetails, on_delete=models.CASCADE)
    customer=models.ForeignKey(Customer,on_delete=models.CASCADE,null=True)
    title=models.TextField(max_length=200)
    document=models.FileField(upload_to='doc/')
    
class CustomerContactPersons(models.Model):
    login_details = models.ForeignKey(LoginDetails, on_delete=models.CASCADE,null=True,blank=True)
    company = models.ForeignKey(CompanyDetails,on_delete=models.CASCADE,null=True,blank=True)
    customer = models.ForeignKey(Customer,on_delete=models.CASCADE,null=True,blank=True)

    title = models.CharField(max_length=220,null=True,blank=True)
    first_name = models.CharField(max_length=220,null=True,blank=True)
    last_name = models.CharField(max_length=220,null=True,blank=True)
    email = models.EmailField(max_length=220,null=True,blank=True)
    work_phone = models.CharField(max_length=220,null=True,blank=True)
    mobile = models.CharField(max_length=220,null=True,blank=True)
    skype = models.CharField(max_length=220,null=True,blank=True)
    designation = models.CharField(max_length=220,null=True,blank=True)
    department = models.CharField(max_length=220,null=True,blank=True)


class CustomerHistory(models.Model):
    login_details = models.ForeignKey(LoginDetails, on_delete=models.CASCADE,null=True,blank=True)
    company = models.ForeignKey(CompanyDetails,on_delete=models.CASCADE,null=True,blank=True)
    customer = models.ForeignKey(Customer,on_delete=models.CASCADE,null=True,blank=True)

    action = models.CharField(max_length=220,null=True,blank=True)
    date = models.DateField(auto_now_add=True, null=True, blank=True)
    
class BankAccount(models.Model):
    login_details = models.ForeignKey(LoginDetails, on_delete=models.CASCADE,null=True,blank=True)
    company=models.ForeignKey(CompanyDetails,on_delete=models.CASCADE,null=True,blank=True)
    bank=models.ForeignKey(Banking, on_delete=models.CASCADE,null=True,blank=True)
    customer_name = models.CharField(max_length=220,null=True)
    alias = models.CharField(max_length=100,null=True)
    phone_number = models.CharField(max_length=10,null=True)
    email = models.EmailField(max_length=100,null=True)
    account_type = models.CharField(max_length=100,null=True)
    bankname=models.CharField(max_length=100,null=True)
    account_number = models.CharField(max_length=15,null=True)
    ifsc_code = models.CharField(max_length=100,null=True)
    swift_code = models.CharField(max_length=100,null=True)
    branch_name = models.CharField(max_length=100,null=True)
    cheque_book_range = models.CharField(max_length=100,null=True)
    enable_cheque_printing = models.CharField(max_length=100,null=True)
    cheque_printing_configuration = models.CharField(max_length=100,null=True)
    mailing_name = models.CharField(max_length=100,null=True)
    address = models.TextField(max_length=100,null=True)
    country = models.CharField(max_length=100,null=True)
    state = models.CharField(max_length=100,null=True)
    pin = models.CharField(max_length=100,null=True)
    pan_number = models.CharField(max_length=100,null=True)
    registration_type = models.CharField(max_length=100,null=True)
    gst_num = models.CharField(max_length=100,null=True)
    alter_gst_details = models.CharField(max_length=100,null=True)
    date = models.DateField(auto_now_add=True, null=True)
    amount_type = models.CharField(max_length=100,null=True)
    amount = models.DecimalField(max_digits=10, decimal_places=2,default=0.00)
    status=models.CharField(max_length=10,default='Active',null=True)
    
    
class BankAccountHistory(models.Model):
    company=models.ForeignKey(CompanyDetails, on_delete=models.CASCADE,null=True,blank=True)
    logindetails= models.ForeignKey(LoginDetails, on_delete=models.CASCADE,null=True,blank=True)
    bank_holder=models.ForeignKey(BankAccount, on_delete=models.CASCADE,null=True,blank=True)
    date = models.DateField(auto_now_add=True, null=True)
    action = models.CharField(max_length=100,null=True)
    
class Loan_Term(models.Model):
    duration= models.IntegerField(null=True,blank=True)
    term = models.CharField(max_length=255,null=True,blank=True)
    days = models.IntegerField(null=True,blank=True)
    company = models.ForeignKey(CompanyDetails,on_delete=models.CASCADE,null=True,blank=True)
    
    
class loan_account(models.Model):
    bank_holder=models.ForeignKey(BankAccount,on_delete=models.CASCADE,null=True)
    logindetails = models.ForeignKey(LoginDetails, on_delete=models.CASCADE,null=True,blank=True)
    company = models.ForeignKey(CompanyDetails, on_delete=models.CASCADE,null=True,blank=True)
    loan_term = models.ForeignKey(Loan_Term, on_delete=models.CASCADE,null=True,blank=True)
    account_number = models.CharField(max_length=15, unique=True,null=True)
    loan_amount=models.IntegerField()
    balance=models.IntegerField(default=0,null=True)
    lender_bank=models.CharField(max_length=255)
    loan_date = models.DateField()
    payment_method=models.CharField(max_length=255)
    upi_id=models.CharField(max_length=255,default='', null=True, blank=True)
    cheque=models.CharField(max_length=255,default='', null=True, blank=True)
    payment_accountnumber=models.CharField(max_length=255,default='', null=True, blank=True)
    processing_method=models.CharField(max_length=255)
    processing_upi=models.CharField(max_length=255,default='', null=True, blank=True)
    processing_cheque=models.CharField(max_length=255,default='', null=True, blank=True)
    processing_acc=models.CharField(max_length=255,default='', null=True, blank=True)
    processing_fee=models.IntegerField(default='', null=True, blank=True)
    term=models.CharField(max_length=15,default='', null=True, blank=True)
    interest=models.IntegerField(default='', null=True, blank=True)
    description=models.CharField(max_length=255,default='', null=True, blank=True)
    status= models.TextField(default='Active')
    
    
class LoanRepayemnt(models.Model):
    login_details = models.ForeignKey(LoginDetails, on_delete=models.CASCADE,null=True,blank=True)
    company=models.ForeignKey(CompanyDetails,on_delete=models.CASCADE,null=True,blank=True)
    loan=models.ForeignKey(loan_account, on_delete=models.CASCADE,null=True,blank=True)
    principal_amount=models.IntegerField(null=True,blank=True)
    interest_amount=models.IntegerField(null=True,blank=True)
    payment_method=models.CharField(max_length=255)
    upi_id=models.CharField(max_length=255,default=None, null=True, blank=True)
    cheque=models.CharField(max_length=255,default=None, null=True, blank=True)
    account_number=models.CharField(max_length=255,default=None, null=True, blank=True)
    payment_date=models.DateField(default=date.today)
    total_amount=models.IntegerField(null=True,blank=True)
    type=models.CharField(max_length=255,null=True)
    
class LoanAccountHistory(models.Model):
    login_details = models.ForeignKey(LoginDetails, on_delete=models.CASCADE,null=True,blank=True)
    company=models.ForeignKey(CompanyDetails,on_delete=models.CASCADE,null=True,blank=True)
    loan=models.ForeignKey(loan_account,on_delete=models.CASCADE,null=True,blank=True)
    date=models.DateField(default=date.today)
    action=models.CharField(max_length=255)
    
    
class LoanRepaymentHistory(models.Model):
    login_details = models.ForeignKey(LoginDetails, on_delete=models.CASCADE,null=True,blank=True)
    company=models.ForeignKey(CompanyDetails,on_delete=models.CASCADE,null=True,blank=True)
    repayment=models.ForeignKey(LoanRepayemnt,on_delete=models.CASCADE,null=True,blank=True)
    date=models.DateField(default=date.today)
    action=models.CharField(max_length=255)
    
    
class Comments(models.Model):
    login_details = models.ForeignKey(LoginDetails, on_delete=models.CASCADE,null=True,blank=True)
    company=models.ForeignKey(CompanyDetails,on_delete=models.CASCADE,null=True,blank=True)
    loan=models.ForeignKey(loan_account,on_delete=models.CASCADE,null=True,blank=True)
    comment=models.CharField(max_length=255)
    
    
class Godown(models.Model):
    date = models.DateField()
    item = models.ForeignKey(Items, on_delete=models.CASCADE,null=True,blank=True)
    hsn = models.CharField(max_length = 250)
    stock_in_hand = models.IntegerField()
    godown_name = models.CharField(max_length = 250)
    godown_address = models.CharField(max_length = 300)
    stock_keeping = models.IntegerField()
    distance = models.IntegerField()
    company = models.ForeignKey(CompanyDetails, on_delete=models.CASCADE,null=True,blank=True)
    login_details = models.ForeignKey(LoginDetails, on_delete=models.CASCADE,null=True,blank=True)
    status = models.CharField(max_length=200, default = 'Active', null=True)
    action = models.CharField(max_length=200, null=True)
    file = models.FileField(upload_to='file/', null=True, blank=True)


class GodownHistory(models.Model):
    company = models.ForeignKey(CompanyDetails, on_delete=models.CASCADE,null=True,blank=True)
    login_details = models.ForeignKey(LoginDetails, on_delete=models.CASCADE,null=True,blank=True)
    godown = models.ForeignKey(Godown, on_delete=models.CASCADE,null=True,blank=True)
    date = models.DateField()
    action = models.CharField(max_length = 250)

class GodownComments(models.Model):
    company = models.ForeignKey(CompanyDetails, on_delete=models.CASCADE,null=True,blank=True)
    login_details = models.ForeignKey(LoginDetails, on_delete=models.CASCADE,null=True,blank=True)
    godown = models.ForeignKey(Godown, on_delete=models.CASCADE,null=True,blank=True)
    comment = models.CharField(max_length = 250)
    
class Holiday_history(models.Model):
    company=models.ForeignKey(CompanyDetails, on_delete=models.CASCADE, null=True,blank=True)
    user = models.ForeignKey(LoginDetails, on_delete=models.CASCADE, null=True, blank=True)
    holiday = models.ForeignKey(Holiday, on_delete=models.CASCADE, null=True, blank=True)
    date = models.DateField(null=True, blank=True)
    action = models.CharField(max_length=255, null=True, blank=True)
    
class Comment_holiday(models.Model):
    holiday_details = models.ForeignKey(Holiday, on_delete = models.CASCADE,null=True,blank=True)
    comment = models.CharField(max_length=255, null=True, blank=True)
    user = models.ForeignKey(LoginDetails, on_delete=models.CASCADE, null=True, blank=True)
    company=models.ForeignKey(CompanyDetails, on_delete=models.CASCADE, null=True,blank=True)
    
    
class JournalRecievedIdModel(models.Model):
    user = models.CharField(max_length=200,null=True,blank=True)
    company = models.ForeignKey(CompanyDetails, on_delete=models.CASCADE, null=True, blank=True)
    staff = models.ForeignKey(StaffDetails, on_delete=models.CASCADE, null=True, blank=True)
    pattern = models.CharField(max_length=255,null=True)
    ref_number = models.CharField(max_length=255,null=True)
    jn_rec_number = models.CharField(max_length=255,null=True) 

      
class Journal(models.Model):
    STATUS_CHOICES = (
        ('draft', 'Draft'),
        ('save', 'Save'),
    )
    company = models.ForeignKey(CompanyDetails, on_delete=models.CASCADE, null=True, blank=True)
    staff = models.ForeignKey(StaffDetails, on_delete=models.CASCADE, null=True, blank=True)
    login_details = models.ForeignKey(LoginDetails, on_delete=models.CASCADE, null=True, blank=True)
    date = models.DateField(null=True,auto_now_add=True)
    journal_no = models.CharField(max_length=255,null=True)  
    reference_no = models.IntegerField(null=True)
    notes = models.TextField(blank=True,null=True)
    currency = models.CharField(max_length=255,null=True)
    journal_type = models.CharField(max_length=255,null=True)
    attachment = models.FileField(upload_to='journal_attachments/', blank=True,null=True)
    total_debit = models.DecimalField(max_digits=10, decimal_places=2, default=0,null=True)
    total_credit = models.DecimalField(max_digits=10, decimal_places=2, default=0,null=True)
    debit_difference = models.DecimalField(max_digits=10, decimal_places=2, default=0,null=True)
    credit_difference = models.DecimalField(max_digits=10, decimal_places=2, default=0,null=True)
    status = models.CharField(max_length=255, choices=STATUS_CHOICES, null=True, blank=True)
     
    def __str__(self):
        return self.journal_no
    
class JournalEntry(models.Model):
    journal = models.ForeignKey(Journal, on_delete=models.CASCADE,null=True)
    company = models.ForeignKey(CompanyDetails, on_delete=models.CASCADE, null=True, blank=True)
    login = models.ForeignKey(LoginDetails, on_delete=models.CASCADE, null=True, blank=True)
    staff = models.ForeignKey(StaffDetails, on_delete=models.CASCADE, null=True, blank=True)
    account = models.CharField(max_length=200,null=True)
    description = models.TextField(null=True)
    contact = models.CharField(max_length=200,null=True)
    debits = models.DecimalField(max_digits=10, decimal_places=2,null=True)
    credits = models.DecimalField(max_digits=10, decimal_places=2,null=True)
    
class JournalTransactionHistory(models.Model):
    company = models.ForeignKey(CompanyDetails, on_delete=models.CASCADE, null=True, blank=True)
    login_details = models.ForeignKey(LoginDetails, on_delete=models.CASCADE, null=True, blank=True)
    staff = models.ForeignKey(StaffDetails, on_delete=models.CASCADE, null=True, blank=True)
    journal = models.ForeignKey(Journal, on_delete=models.CASCADE)
    date = models.DateField(auto_now_add=True, null=True)
    action_choices = [
        ('Created', 'Created'), 
        ('Edited', 'Edited')
    ]
    action = models.CharField(max_length=10, choices=action_choices, null=True)

class JournalComment(models.Model):
    company = models.ForeignKey(CompanyDetails, on_delete=models.CASCADE, null=True, blank=True)
    login_details = models.ForeignKey(LoginDetails, on_delete=models.CASCADE, null=True, blank=True)
    staff = models.ForeignKey(StaffDetails, on_delete=models.CASCADE, null=True, blank=True)
    journal = models.ForeignKey(Journal, on_delete=models.CASCADE)
    comment = models.TextField()
    date_added = models.DateField(auto_now_add=True)  
    
    
#-----------------invoice -----------------------------#
class invoice(models.Model):
    company = models.ForeignKey(CompanyDetails,on_delete=models.CASCADE,null=True,blank=True)
    login_details = models.ForeignKey(LoginDetails, on_delete=models.CASCADE,null=True,blank=True)
    customer = models.ForeignKey(Customer,on_delete=models.CASCADE,null=True,blank=True)
    payment_terms = models.ForeignKey(Company_Payment_Term,on_delete=models.CASCADE,null=True,blank=True)

    customer_email=models.EmailField(max_length=220,null=True,blank=True)
    customer_billingaddress=models.CharField(max_length=220,null=True,blank=True)
    customer_GSTtype=models.CharField(max_length=220,null=True,blank=True)
    customer_GSTnumber=models.CharField(max_length=220,null=True,blank=True)
    customer_place_of_supply=models.CharField(max_length=220,null=True,blank=True)
    date = models.DateField(auto_now_add=True, null=True, blank=True) 
    expiration_date = models.DateField(auto_now_add=True, null=True, blank=True) 
    reference_number=models.IntegerField(blank=True,null=True,)
    invoice_number=models.CharField(max_length=220,null=True,blank=True) 
    payment_method=models.CharField(max_length=220,null=True,blank=True) 
    cheque_number=models.CharField(max_length=220,null=True,blank=True) 
    UPI_number=models.CharField(max_length=220,null=True,blank=True) 
    bank_account_number=models.CharField(max_length=220,null=True,blank=True) 
    description=models.CharField(max_length=220,null=True,blank=True) 
    terms_and_condition=models.CharField(max_length=220,null=True,blank=True) 
    document=models.FileField(upload_to="images/",null=True)
    sub_total=models.FloatField(default=0.0, null=True, blank=True)
    CGST=models.FloatField(default=0.0, null=True, blank=True)
    SGST=models.FloatField(default=0.0, null=True, blank=True)
    IGST = models.FloatField(default=0.0, null=True, blank=True)

    tax_amount=models.FloatField(default=0.0, null=True, blank=True)
    shipping_charge=models.FloatField(default=0.0, null=True, blank=True)
    adjustment=models.FloatField(default=0.0, null=True, blank=True)
    grand_total=models.FloatField(default=0.0, null=True, blank=True)
    advanced_paid=models.FloatField(default=0.0, null=True, blank=True)
    balance=models.FloatField(default=0.0, null=True, blank=True)
    status=models.CharField(max_length=220,null=True,blank=True)  

class invoiceHistory(models.Model):
    login_details = models.ForeignKey(LoginDetails, on_delete=models.CASCADE,null=True,blank=True)
    company = models.ForeignKey(CompanyDetails,on_delete=models.CASCADE,null=True,blank=True)
    invoice = models.ForeignKey(invoice,on_delete=models.CASCADE,null=True,blank=True)

    action = models.CharField(max_length=220,null=True,blank=True)
    date = models.DateField(auto_now_add=True, null=True, blank=True) 
    
class invoiceReference(models.Model):
    reference_number = models.CharField(max_length=220,null=True,blank=True)
    company = models.ForeignKey(CompanyDetails,on_delete=models.CASCADE,null=True,blank=True)
    staff = models.ForeignKey(StaffDetails,on_delete=models.CASCADE,null=True,blank=True)


class invoiceitems(models.Model):
    company = models.ForeignKey(CompanyDetails,on_delete=models.CASCADE)
    logindetails = models.ForeignKey(LoginDetails,on_delete=models.CASCADE)
    Items = models.ForeignKey(Items,on_delete=models.CASCADE)
    invoice = models.ForeignKey(invoice,on_delete=models.CASCADE,null=True,blank=True)

    hsn = models.CharField(max_length=220,null=True,blank=True)
    quantity=models.IntegerField(blank=True,null=True,default=0)
    price=models.FloatField(default=0.0, null=True, blank=True)
    tax_rate=models.FloatField(default=0.0, null=True, blank=True)
    discount=models.FloatField(default=0.0, null=True, blank=True)
    total=models.FloatField(default=0.0, null=True, blank=True)

class invoicecomments(models.Model):
    company = models.ForeignKey(CompanyDetails,on_delete=models.CASCADE)
    invoice = models.ForeignKey(invoice,on_delete=models.CASCADE,null=True,blank=True)
    comments = models.CharField(max_length=500,null=True,blank=True)

#End