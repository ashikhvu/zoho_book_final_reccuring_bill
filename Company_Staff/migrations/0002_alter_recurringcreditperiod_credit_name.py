# Generated by Django 3.2.23 on 2024-03-26 06:50

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('Company_Staff', '0001_initial'),
    ]

    operations = [
        migrations.AlterField(
            model_name='recurringcreditperiod',
            name='credit_name',
            field=models.CharField(blank=True, max_length=255, null=True),
        ),
    ]
