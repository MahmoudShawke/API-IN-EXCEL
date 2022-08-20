from flask import Flask, request

import pandas as pd

app = Flask(__name__)


def save_data_to_excel(excel_name, sheet_name, data):
    columns = []
    for k, v in data.items():
        columns.append(k)
    with pd.ExcelWriter(excel_name) as writer:
        df = pd.DataFrame(data, index=None)
        df.to_excel(writer, sheet_name=sheet_name, index=False, columns=columns)


@app.route('/get/<int:id>')
def get(id=0):
    if id == 0:
        excel_data_df = pd.read_excel('testing.xlsx', usecols=["الترتيب", "الرواية", "المؤلف", "البلد"])
        dataJson = excel_data_df.to_json(orient='records')
        return dataJson.encode('UTF-8').decode('unicode_escape')
    else:
        df = pd.read_excel('testing.xlsx', usecols=["الترتيب", "الرواية", "المؤلف", "البلد"])
        df.set_index("الترتيب", inplace=True)
        df = df.loc[[id]]
        dataJson = df.to_json(orient='records')
        return dataJson.encode('UTF-8').decode('unicode_escape')


@app.route('/insert', methods=["POST"])
def insert():
    excel_data_df = pd.read_excel('testing.xlsx', usecols=["الترتيب", "الرواية", "المؤلف", "البلد"])

    value = {
        "الترتيب": request.args.get('nov'),
        "الرواية": request.args.get('novel'),
        "المؤلف": request.args.get('author'),
        "البلد": request.args.get('country')
    }
    data = excel_data_df.append(value, ignore_index=True)
    save_data_to_excel(excel_name="testing.xlsx", sheet_name='test', data=data)
    return ['true']


@app.route('/update/<int:id>', methods=["POST"])
def update(id):
    if request.method == 'POST':
        df = pd.read_excel('testing.xlsx', usecols=["الترتيب", "الرواية", "المؤلف", "البلد"])
        df.loc[df['الترتيب'] == id, 'الترتيب'] = request.args.get('nov')
        df.loc[df['الترتيب'] == id, 'الرواية'] = request.args.get('novel')
        df.loc[df['الترتيب'] == id, 'المؤلف'] = request.args.get('author')
        df.loc[df['الترتيب'] == id, 'البلد'] = request.args.get('country')
        df.to_excel('testing.xlsx', index=False)

    return ['true']


@app.route('/delete/<int:id>', methods=["Delete"])
def delete(id):
    df = pd.read_excel('testing.xlsx', usecols=["الترتيب", "الرواية", "المؤلف", "البلد"])
    df.drop(df[df['الترتيب'] == id].index, inplace=True)
    df.to_excel('testing.xlsx', index=False)
    return ['true']


if __name__ == '__main__':
    app.run()
