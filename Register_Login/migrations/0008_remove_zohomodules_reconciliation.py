# Generated by Django 4.2 on 2023-12-30 07:20

from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('Register_Login', '0007_companydetails_position'),
    ]

    operations = [
        migrations.RemoveField(
            model_name='zohomodules',
            name='reconciliation',
        ),
    ]
