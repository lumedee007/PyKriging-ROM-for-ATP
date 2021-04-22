# This code is a function that predicts the 42 products of the pyrolysis products based on the inoputs in aspen 
def pyrolysis_products():
    import pandas as pd
    import pickle
    import xlwings as xw
    
# The next line imports the ranzi data to extract the products 
    ranzi_data = r"C:\Users\olumideo\Desktop\Research\Ranzi full component\Ranzi_data_42prod.xlsx"
    df = pd.read_excel(ranzi_data)
    headers = df.keys().values.tolist()
# The next line creates a list of the products by removing the inputs from the header's list     
    products = headers[5:]
# The next line creates an empty list, the products predicted are going to be stored in this list while we loop through each product model    
    predictions =[]
# The next lines calls the active excel workbook opened by aspen plus.   
    wb =xw.books.active
    sht=wb.sheets['products']
    sht1=wb.sheets['feeds']
# The for loop will loop through the products models that was trained and saved earlier, predict and append each product to the prediction list 
    for name in products:
        model = 'model' + name + ".sav"
        model2 = pickle.load(open(model,'rb'))
        pred = model2.predict(sht1.range('F3:J3').value)
        predictions.append(pred)
        sht.range('C3').value = sht1.range('C3','E3').value
   
    
# The next line will return the predictions to the excel workbook. 
    sht.range('F3').value = predictions
