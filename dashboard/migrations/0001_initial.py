# Generated by Django 3.0.1 on 2020-01-10 11:35

from django.db import migrations, models


class Migration(migrations.Migration):

    initial = True

    dependencies = [
    ]

    operations = [
        migrations.CreateModel(
            name='Sale_detail',
            fields=[
                ('id', models.AutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('customer', models.CharField(max_length=100)),
                ('cust_c_loc', models.CharField(max_length=100)),
                ('cust_j_loc', models.CharField(max_length=100)),
            ],
        ),
    ]
