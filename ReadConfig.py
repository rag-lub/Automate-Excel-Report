import csv
def Init_Inputs(config_inputs): 
    with open("config.csv","r") as csv_file:
        csv_reader = csv.DictReader(csv_file,delimiter="=")
        for line in csv_reader:
            k = line["variable"]
            if k.startswith("#"): #skips comment rows in csv config file
                continue
            elif k == "Geo":
                config_inputs[k] = eval("{"+line["value"]+"}")
            else:
                config_inputs[k] = line["value"]        
    return config_inputs
