# Generated by Django 3.2.23 on 2024-04-01 04:03

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('Company_Staff', '0008_rename_text_amount_recurring_bills_tax_amount'),
    ]

    operations = [
        migrations.AddField(
            model_name='recurring_bills',
            name='repeat_every_duration',
            field=models.CharField(blank=True, max_length=255, null=True),
        ),
        migrations.AddField(
            model_name='recurring_bills',
            name='repeat_every_type',
            field=models.IntegerField(blank=True, null=True),
        ),
    ]
