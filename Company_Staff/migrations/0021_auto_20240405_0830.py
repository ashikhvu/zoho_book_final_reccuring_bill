# Generated by Django 3.2.23 on 2024-04-05 08:30

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('Company_Staff', '0020_auto_20240405_0830'),
    ]

    operations = [
        migrations.AlterField(
            model_name='recurritemslist',
            name='taxGST',
            field=models.CharField(blank=True, max_length=255, null=True),
        ),
        migrations.AlterField(
            model_name='recurritemslist',
            name='taxIGST',
            field=models.CharField(blank=True, max_length=255, null=True),
        ),
    ]