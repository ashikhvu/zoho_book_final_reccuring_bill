# Generated by Django 3.2.23 on 2024-03-25 13:21

from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('Register_Login', '0017_alter_trialperiod_interested_in_buying'),
        ('Company_Staff', '0026_auto_20240325_1301'),
    ]

    operations = [
        migrations.RenameModel(
            old_name='RepeatEvery',
            new_name='ReccuringRepeatEvery',
        ),
    ]
