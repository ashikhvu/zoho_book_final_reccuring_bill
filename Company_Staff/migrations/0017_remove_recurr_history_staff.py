# Generated by Django 3.2.23 on 2024-04-04 17:20

from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('Company_Staff', '0016_recurr_history'),
    ]

    operations = [
        migrations.RemoveField(
            model_name='recurr_history',
            name='staff',
        ),
    ]