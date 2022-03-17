import pymongo
import csv
import pandas as pd


def get_dataset_list():
    with open('covid_19_india.csv') as f:
        dataset = [{k:v for k, v in row.items()} for row in csv.DictReader(f, skipinitialspace = True)]
    return dataset


def insert_data():
    dataset = get_dataset_list()
    try:
        client = pymongo.MongoClient('mongo connection string')
        db = client['covid']
        collection = db['covid_state_dataset']
        collection.insert_many(dataset)
        return "Data inserted successfully"
    except Exception as e:
        print(e)
        return "Sorry! something went wrong."

def create_xlsx_workbook():
    try:
        # connection with mogodb
        client = pymongo.MongoClient('monogo connection string')
        db = client['covid']
        collection = db['covid_state_dataset']
        # state list for which we want the dataset
        state_list = ["Kerala", "Bihar", "Uttar Pradesh"]
        sheet_list= {}
        for state in state_list:
            # getting records of each state
            dataset = list(collection.find({"State/UnionTerritory": state}))
            # converting the dataset into dataframe
            dataset = pd.DataFrame.from_dict(dataset)
            # creating a dictionaries with key as state name and value as the dataframe
            sheet_list[state] = dataset
        writer = pd.ExcelWriter('./covid.xlsx', engine='xlsxwriter')
        # looping through the state names and writing their dataframe into sheets
        for sheet_name in sheet_list.keys():
            sheet_list[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)
        writer.save()
        return "Workbook created successfully"

    except Exception as e:
        print(e)
        return "Sorry! something went wrong."

print(create_xlsx_workbook())


