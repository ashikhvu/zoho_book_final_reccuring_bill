# Generated by Django 4.2 on 2023-12-21 06:33

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('Register_Login', '0002_logindetails_distributor_id_and_more'),
    ]

    operations = [
        migrations.AlterField(
            model_name='logindetails',
            name='distributor_id',
            field=models.CharField(blank=True, default='', max_length=100, null=True),
        ),
    ]
