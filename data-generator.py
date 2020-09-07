# -*- coding: utf-8 -*-
"""
Creates a randomly generated dataset of vehicles and attributes. 
Takes input for path and number of records from console.

"""
import random
import datetime
import pandas

max_records = int(input("How many records would you like to create?  "))
save_path = input("Where would you like this file saved?  ")

row_count = 0
mark_up_percent = 30
model = ['sedan','coupe','minivan','pickup','suv','crossover','cargo_van','sports_car','motorcycle']
make = ['ford','chevrolet','cadillac','honda','nissan','mercedes_benz','toyota', 'bmw','dodge','chrysler']
cylinders = ['I4','I6','V6','V8']
fuel = ['diesel','gas','electric']
trim = ['basic','touring','luxury','sport','special_edition']
color = ['white','black','gray','silver','gold','blue','red','yellow','green','brown','orange']
condition = ['new','like_new','good','fair','poor']
transmission = ['manual','automatic']
cars_data = []



def randomDate():
    list_date_start = datetime.date(2017, 12, 31)
    list_date_end = datetime.date(2020, 9, 1)
    list_time_between = list_date_end - list_date_start
    list_days_between = list_time_between.days
    random_number_days = random.randrange(list_days_between)
    return list_date_start + datetime.timedelta(days=random_number_days)

class Car:
    def __init__(self):
        self.model = random.choice(model)
        self.make = random.choice(make)
        self.cylinders = random.choice(cylinders)
        self.fuel = random.choice(fuel)
        self.trim = random.choice(trim)
        self.color = random.choice(color)
        self.year = random.randrange(2000,2020)
        self.miles = random.randrange(25,300000)
        self.blue_book_val = random.randrange(3000, 50000)
        self.list_price = round(self.blue_book_val * (1 + (mark_up_percent/100)), 2)
        self.condition = random.choice(condition)
        self.transmission = random.choice(transmission)
        self.list_date = randomDate()



while row_count < max_records:
    car = Car()
    cars_data.append(
        [car.make,car.model,car.cylinders,car.fuel,
        car.trim,car.color,car.year,car.miles,
        car.blue_book_val,car.list_price,car.condition,car.transmission,
        car.list_date]
        )
    row_count = row_count+1
    
df = pandas.DataFrame(cars_data,
    columns=[
        'make','model','cylinders','fuel',
        'trim','color','year','miles',
        'blue_book_val','list_price', 'condition','transmission',
        'list_date'
        ]
    )

df.to_csv(path_or_buf=save_path)
