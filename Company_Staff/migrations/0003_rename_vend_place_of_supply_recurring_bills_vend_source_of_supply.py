# Generated by Django 3.2.23 on 2024-03-30 07:21

from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('Company_Staff', '0002_alter_recurringcreditperiod_credit_name'),
    ]

    operations = [
        migrations.RenameField(
            model_name='recurring_bills',
            old_name='vend_place_of_supply',
            new_name='vend_source_of_supply',
        ),
    ]
