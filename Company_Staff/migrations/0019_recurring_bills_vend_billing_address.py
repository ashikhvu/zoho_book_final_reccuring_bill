# Generated by Django 3.2.23 on 2024-04-05 07:57

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('Company_Staff', '0018_rename_check_no_recurring_bills_cheque_no'),
    ]

    operations = [
        migrations.AddField(
            model_name='recurring_bills',
            name='vend_billing_address',
            field=models.TextField(blank=True, null=True),
        ),
    ]
