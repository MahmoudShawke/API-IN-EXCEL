from flask import Flask

import pandas as pd


app = Flask(__name__)

def save_data_to_excel(excel_name, sheet_name, data):
    columns = []
    for k, v in data.items():
        columns.append(k)
    with pd.ExcelWriter(excel_name) as writer:
        df = pd.DataFrame(data, index=None)
        df.to_excel(writer, sheet_name=sheet_name, index=False, columns=columns )



@app.route('/get')
def get():
    excel_data_df = pd.read_excel('testing.xlsx', usecols=["الترتيب", "الرواية", "المؤلف", "البلد"])
    dataJson=excel_data_df.to_json(orient='records')
    return dataJson.encode('UTF-8').decode('unicode_escape')

@app.route('/insert',methods=["POST"])
def insert():
    excel_data_df = pd.read_excel('testing.xlsx', usecols=["الترتيب", "الرواية", "المؤلف", "البلد"])
    value= {
      "الترتيب":110,
      "الرواية":"https.tessting.com",
      "المؤلف":"testing7%D8%A7%D9%86%D9%8A",
      "البلد":"testinghttps:\/\/ar.wikipedia.org\/wiki\/مصر"
   }
    data=excel_data_df.append(value,ignore_index=True)
    save_data_to_excel(excel_name="testing.xlsx", sheet_name='test', data=data)
    return ['true']
@app.route('/update',methods=["POST"])
def update():
    df = pd.read_excel('testing.xlsx', usecols=["الترتيب", "الرواية", "المؤلف", "البلد"])
    df.loc[df['الترتيب']==110 ,'الترتيب']=150
    df.to_excel('testing.xlsx',index=False)

    return ['true']

@app.route('/delete',methods=["Delete"])
def delete():
    df = pd.read_excel('testing.xlsx', usecols=["الترتيب", "الرواية", "المؤلف", "البلد"])
    df.drop(df[df['الترتيب']==110 ].index,inplace=True)
    df.to_excel('testing.xlsx',index=False)
    return ['true']




if __name__ == '__main__':
    app.run()
