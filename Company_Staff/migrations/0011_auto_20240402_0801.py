# Generated by Django 3.2.23 on 2024-04-02 08:01

from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('Company_Staff', '0010_auto_20240402_0418'),
    ]

    operations = [
        migrations.RenameField(
            model_name='recurringrecievedid',
            old_name='rec_rec_no',
            new_name='recc_rec_number',
        ),
        migrations.RenameField(
            model_name='recurringrecievedid',
            old_name='rec_ref_no',
            new_name='ref_number',
        ),
    ]